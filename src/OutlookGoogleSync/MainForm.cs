using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class MainForm : Form
    {
        internal const string FILENAME = "settings.xml";

        private Timer _ogstimer;
        private DateTime _oldtime;
        private List<int> _minuteOffsets = new List<int>();
        private FormWindowState _previousWindowState = FormWindowState.Normal;

        private Dictionary<Event, string> _googleEventSignatures = new Dictionary<Event, string>();
        private Dictionary<AppointmentItem, string> _outlookAppointmentSignatures = new Dictionary<AppointmentItem, string>();

        public MainForm()
        {
            InitializeComponent();
            labelAbout.Text = labelAbout.Text.Replace("{version}", System.Windows.Forms.Application.ProductVersion);

            //load settings/create settings file
            if (File.Exists(FILENAME))
            {
                Settings.Instance = XMLManager.import<Settings>(FILENAME);
            }
            else
            {
                XMLManager.Export(Settings.Instance, FILENAME);
            }

            //update GUI from Settings
            numericUpDownDaysInThePast.Text = Settings.Instance.DaysInThePast.ToString();
            numericUpDownDaysInTheFuture.Text = Settings.Instance.DaysInTheFuture.ToString();
            textBoxMinuteOffsets.Text = Settings.Instance.MinuteOffsets;
            comboBoxCalendars.Items.Add(Settings.Instance.SelectedGoogleCalendar);
            comboBoxCalendars.SelectedIndex = 0;
            checkBoxSyncEveryHour.Checked = Settings.Instance.SyncEveryHour;
            checkBoxShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
            checkBoxStartInTray.Checked = Settings.Instance.StartInTray;
            checkBoxMinimizeToTray.Checked = Settings.Instance.MinimizeToTray;
            checkBoxAddDescription.Checked = Settings.Instance.AddDescription;
            checkBoxAddAttendees.Checked = Settings.Instance.AddAttendeesToDescription;
            checkBoxAddReminders.Checked = Settings.Instance.AddReminders;
            checkBoxCreateFiles.Checked = Settings.Instance.CreateTextFiles;

            //Start in tray?
            if (checkBoxStartInTray.Checked)
            {
                WindowState = FormWindowState.Minimized;
            }
            else
            {
                ShowInTaskbar = true;
            }

            //set up timer (every 30s) for checking the minute offsets
            _ogstimer = new Timer();
            _ogstimer.Interval = 30000;
            _ogstimer.Tick += new EventHandler(ogstimer_Tick);
            _ogstimer.Start();
            _oldtime = DateTime.Now;

            //set up tooltips for some controls
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 10000;
            toolTip1.InitialDelay = 500;
            toolTip1.ReshowDelay = 200;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(comboBoxCalendars,
                "The Google Calendar to synchonize with.");
            toolTip1.SetToolTip(textBoxMinuteOffsets,
                "One ore more Minute Offsets at which the sync is automatically started each hour. \n" +
                "Separate by comma (e.g. 5,15,25).");
            toolTip1.SetToolTip(checkBoxAddAttendees,
                "While Outlook has fields for Organizer, RequiredAttendees and OptionalAttendees, Google has not.\n" +
                "If checked, this data is added at the end of the description as text.");
            toolTip1.SetToolTip(checkBoxAddReminders,
                "If checked, the reminder set in outlook will be carried over to the Google Calendar entry (as a popup reminder).");
            toolTip1.SetToolTip(checkBoxCreateFiles,
                "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
                "to 4 separate text files in the application's directory (named \"export_*.txt\"). \n" +
                "Only for debug/diagnostic purposes.");
            toolTip1.SetToolTip(checkBoxAddDescription,
                "The description may contain email addresses, which Outlook may complain about (PopUp-Message: \"Allow Access?\" etc.). \n" +
                "Turning this off allows OutlookGoogleSync to run without intervention in this case.");
            toolTip1.SetToolTip(buttonDeleteAll,
                "Delete all calendar items from your Google calendar which were synced from your Outlook calendar in the set time range.");
        }

        private void ogstimer_Tick(object sender, EventArgs e)
        {
            if (!checkBoxSyncEveryHour.Checked)
                return;
            DateTime newtime = DateTime.Now;
            if (newtime.Minute != _oldtime.Minute)
            {
                _oldtime = newtime;
                if (_minuteOffsets.Contains(newtime.Minute))
                {
                    if (checkBoxShowBubbleTooltips.Checked)
                        notifyIcon1.ShowBalloonTip(
                            500,
                            "OutlookGoogleSync",
                            "Sync started at desired minute offset " + newtime.Minute.ToString(),
                            ToolTipIcon.Info
                            );
                    buttonSyncNow_Click(null, null);
                }
            }
        }

        private void buttonGetMyGoogleCalendars_Click(object sender, EventArgs e)
        {
            buttonGetMyCalendars.Enabled = false;
            comboBoxCalendars.Enabled = false;

            try
            {
                GoogleCalendar gcal = new GoogleCalendar();
                List<GoogleCalendarListEntry> calendars = gcal.getCalendars();
                if (calendars != null)
                {
                    comboBoxCalendars.Items.Clear();
                    foreach (GoogleCalendarListEntry mcle in calendars)
                    {
                        comboBoxCalendars.Items.Add(mcle);
                    }
                    comboBoxCalendars.SelectedIndex = 0;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error:\r\n" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            buttonGetMyCalendars.Enabled = true;
            comboBoxCalendars.Enabled = true;
        }

        private void buttonSyncNow_Click(object sender, EventArgs e)
        {
            _googleEventSignatures.Clear();
            _outlookAppointmentSignatures.Clear();

            if (Settings.Instance.SelectedGoogleCalendar.Id == "")
            {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return;
            }

            buttonSyncNow.Enabled = false;
            buttonDeleteAll.Enabled = false;

            textBoxLogs.Clear();

            DateTime SyncStarted = DateTime.Now;

            try
            {
                logboxout("Sync started at " + SyncStarted.ToString());
                logboxout("--------------------------------------------------");

                logboxout("Reading Outlook Calendar Entries...");
                OutlookCalendar ocal = new OutlookCalendar();
                List<AppointmentItem> OutlookEntries = ocal.getCalendarEntriesInRange();
                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_found_in_outlook.txt"))
                    {
                        foreach (AppointmentItem ai in OutlookEntries)
                        {
                            tw.WriteLine(signature(ai));
                        }
                    }
                }
                logboxout("Found " + OutlookEntries.Count + " Outlook Calendar Entries.");
                logboxout("--------------------------------------------------");

                logboxout("Reading Google Calendar Entries...");
                GoogleCalendar gcal = new GoogleCalendar();
                List<Event> GoogleEntries = gcal.getCalendarEntriesInRange();
                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_found_in_google.txt"))
                    {
                        foreach (Event ev in GoogleEntries)
                        {
                            tw.WriteLine(signature(ev));
                        }
                    }
                }
                logboxout("Found " + GoogleEntries.Count + " Google Calendar Entries.");
                logboxout("--------------------------------------------------");

                string accountName = "(Empty)";
                if (ocal != null || !string.IsNullOrEmpty(ocal.AccountName))
                {
                    accountName = ocal.AccountName;
                }
                List<Event> GoogleEntriesToBeDeleted = identifyGoogleEntriesToBeDeleted(OutlookEntries, GoogleEntries, accountName);
                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_to_be_deleted.txt"))
                    {
                        foreach (Event ev in GoogleEntriesToBeDeleted)
                        {
                            tw.WriteLine(signature(ev));
                        }
                    }
                }
                logboxout(GoogleEntriesToBeDeleted.Count + " Google Calendar Entries to be deleted.");

                //OutlookEntriesToBeCreated ...in Google!
                List<AppointmentItem> OutlookEntriesToBeCreated = identifyOutlookEntriesToBeCreated(OutlookEntries, GoogleEntries);
                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_to_be_created.txt"))
                    {
                        foreach (AppointmentItem ai in OutlookEntriesToBeCreated)
                        {
                            tw.WriteLine(signature(ai));
                        }
                    }
                }
                logboxout(OutlookEntriesToBeCreated.Count + " Entries to be created in Google.");
                logboxout("--------------------------------------------------");


                if (GoogleEntriesToBeDeleted.Count > 0)
                {
                    logboxout("Deleting " + GoogleEntriesToBeDeleted.Count + " Google Calendar Entries...");
                    foreach (Event ev in GoogleEntriesToBeDeleted)
                        gcal.deleteCalendarEntry(ev);
                    logboxout("Done.");
                    logboxout("--------------------------------------------------");
                }

                if (OutlookEntriesToBeCreated.Count > 0)
                {
                    logboxout("Creating " + OutlookEntriesToBeCreated.Count + " Entries in Google...");
                    foreach (AppointmentItem ai in OutlookEntriesToBeCreated)
                    {
                        Event ev = new Event();

                        ev.Start = new EventDateTime();
                        ev.End = new EventDateTime();

                        if (ai.AllDayEvent)
                        {
                            ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                            ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(ai.Start);
                            ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(ai.End);
                        }
                        ev.Summary = ai.Subject;
                        if (checkBoxAddDescription.Checked)
                        {
                            try
                            {
                                ev.Description = ai.Body;
                            }
                            catch (System.Exception ex)
                            {
                                string startDt = ai.AllDayEvent ? ai.Start.ToShortDateString() : ai.Start.ToString();
                                string endDt = ai.AllDayEvent ? ai.End.ToShortDateString() : ai.End.ToString();
                                logboxout("Error accessing the body of Outlook item. Body will be empty.\r\n    Subject: [" + ev.Summary + "]\r\n    Start: [" + startDt + "]\r\n    End: [" + endDt + "]\r\n    Error: " + ex.Message);
                            }
                        }
                        ev.Location = ai.Location;


                        //consider the reminder set in Outlook
                        if (checkBoxAddReminders.Checked && ai.ReminderSet)
                        {
                            ev.Reminders = new Event.RemindersData();
                            ev.Reminders.UseDefault = false;
                            EventReminder reminder = new EventReminder();
                            reminder.Method = "popup";
                            reminder.Minutes = ai.ReminderMinutesBeforeStart;
                            ev.Reminders.Overrides = new List<EventReminder>();
                            ev.Reminders.Overrides.Add(reminder);
                        }
                        else
                        {
                            ev.Reminders = new Event.RemindersData();
                            ev.Reminders.UseDefault = false;
                        }

                        if (checkBoxAddAttendees.Checked)
                        {
                            ev.Description += Environment.NewLine;
                            ev.Description += Environment.NewLine + "==============================================";
                            ev.Description += Environment.NewLine + "Added by OutlookGoogleSync (" + accountName + "):" + Environment.NewLine;
                            ev.Description += Environment.NewLine + "ORGANIZER: " + Environment.NewLine + ai.Organizer + Environment.NewLine;
                            ev.Description += Environment.NewLine + "REQUIRED: " + Environment.NewLine + splitAttendees(ai.RequiredAttendees) + Environment.NewLine;
                            if (ai.OptionalAttendees != null)
                            {
                                ev.Description += Environment.NewLine + "OPTIONAL: " + Environment.NewLine + splitAttendees(ai.OptionalAttendees);
                            }
                            ev.Description += Environment.NewLine + "==============================================";
                        }

                        gcal.addEntry(ev);
                    }
                    logboxout("Done.");
                    logboxout("--------------------------------------------------");
                }

                DateTime SyncFinished = DateTime.Now;
                TimeSpan Elapsed = SyncFinished - SyncStarted;
                logboxout("Sync finished at " + SyncFinished.ToString());
                logboxout("Time needed: " + Elapsed.Minutes + " min " + Elapsed.Seconds + " s");
            }
            catch (System.Exception ex)
            {
                logboxout("Error Syncing:\r\n" + ex.ToString());
            }

            buttonSyncNow.Enabled = true;
            buttonDeleteAll.Enabled = true;
        }

        //one attendee per line
        private string splitAttendees(string attendees)
        {
            if (attendees == null)
                return "";
            string[] tmp1 = attendees.Split(';');
            for (int i = 0; i < tmp1.Length; i++)
                tmp1[i] = tmp1[i].Trim();
            return String.Join(Environment.NewLine, tmp1);
        }

        private List<Event> identifyGoogleEntriesToBeDeleted(List<AppointmentItem> outlook, List<Event> google, string accountName)
        {
            List<Event> result = new List<Event>();
            foreach (Event g in google)
            {
                bool found = false;
                foreach (AppointmentItem o in outlook)
                {
                    if (signature(g) == signature(o))
                        found = true;
                }
                if (!found &&
                    !string.IsNullOrEmpty(g.Description) &&
                    g.Description.Contains("Added by OutlookGoogleSync (" + accountName + "):"))
                    result.Add(g);
            }
            return result;
        }

        private List<AppointmentItem> identifyOutlookEntriesToBeCreated(List<AppointmentItem> outlook, List<Event> google)
        {
            List<AppointmentItem> result = new List<AppointmentItem>();
            foreach (AppointmentItem o in outlook)
            {
                bool found = false;
                foreach (Event g in google)
                {
                    if (signature(g) == signature(o))
                        found = true;
                }
                if (!found)
                    result.Add(o);
            }
            return result;
        }

        //creates a standardized summary string with the key attributes of a calendar entry for comparison
        private string signature(AppointmentItem ai)
        {
            if (!_outlookAppointmentSignatures.ContainsKey(ai))
            {
                _outlookAppointmentSignatures.Add(ai, (GoogleCalendar.GoogleTimeFrom(ai.Start) + ";" + GoogleCalendar.GoogleTimeFrom(ai.End) + ";" + ai.Subject + ";" + ai.Location).Trim());
            }

            return _outlookAppointmentSignatures[ai];
        }

        private string signature(Event ev)
        {
            if (!_googleEventSignatures.ContainsKey(ev))
            {
                if (ev.Start != null)
                {
                    if (ev.Start.DateTime == null)
                    {
                        ev.Start.DateTime = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.Start.Date));
                    }
                    if (ev.End.DateTime == null)
                    {
                        ev.End.DateTime = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(ev.End.Date));
                    }
                    _googleEventSignatures.Add(ev, (ev.Start.DateTime + ";" + ev.End.DateTime + ";" + ev.Summary + ";" + ev.Location).Trim());
                }
                else
                {
                    _googleEventSignatures.Add(ev, (ev.Status + ";" + ev.Summary + ";" + ev.Location).Trim());
                }
            }

            return _googleEventSignatures[ev];
        }

        private void logboxout(string s)
        {
            textBoxLogs.Text += s + Environment.NewLine;
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            try
            {
                XMLManager.Export(Settings.Instance, FILENAME);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error Saving:\r\n" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBoxCalendars_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Instance.SelectedGoogleCalendar = (GoogleCalendarListEntry)comboBoxCalendars.SelectedItem;
        }

        private void numericUpDownDaysInThePast_ValueChanged(object sender, EventArgs e)
        {
            Settings.Instance.DaysInThePast = int.Parse(numericUpDownDaysInThePast.Text);
        }

        private void numericUpDownDaysInTheFuture_ValueChanged(object sender, EventArgs e)
        {
            Settings.Instance.DaysInTheFuture = int.Parse(numericUpDownDaysInTheFuture.Text);
        }

        private void textBoxMinuteOffsets_TextChanged(object sender, EventArgs e)
        {
            Settings.Instance.MinuteOffsets = textBoxMinuteOffsets.Text;

            _minuteOffsets.Clear();
            char[] delimiters = { ' ', ',', '.', ':', ';' };
            string[] chunks = textBoxMinuteOffsets.Text.Split(delimiters);
            foreach (string c in chunks)
            {
                int min = 0;
                int.TryParse(c, out min);
                _minuteOffsets.Add(min);
            }
        }


        private void checkBoxSyncEveryHour_CheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.SyncEveryHour = checkBoxSyncEveryHour.Checked;
        }

        private void checkBoxShowBubbleTooltips_CheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.ShowBubbleTooltipWhenSyncing = checkBoxShowBubbleTooltips.Checked;
        }

        private void checkBoxStartInTray_CheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.StartInTray = checkBoxStartInTray.Checked;
        }

        private void checkBoxMinimizeToTray_CheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.MinimizeToTray = checkBoxMinimizeToTray.Checked;
        }

        private void checkBoxAddDescription_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.AddDescription = checkBoxAddDescription.Checked;
        }

        private void checkBoxAddReminders_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.AddReminders = checkBoxAddReminders.Checked;
        }

        private void cbAddAttendees_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.AddAttendeesToDescription = checkBoxAddAttendees.Checked;
        }

        private void cbCreateFiles_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.CreateTextFiles = checkBoxCreateFiles.Checked;
        }

        private void notifyIcon1_Click(object sender, EventArgs e)
        {
            Show();
            WindowState = _previousWindowState;
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (!checkBoxMinimizeToTray.Checked)
            {
                return;
            }

            if (WindowState == FormWindowState.Minimized)
            {
                notifyIcon1.Visible = true;
                ShowInTaskbar = false;
                Hide();
            }
            else
            {
                _previousWindowState = WindowState;
                notifyIcon1.Visible = false;
                ShowInTaskbar = true;
            }
        }

        private void linkLabelUrl_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(linkLabelUrl.Text);
        }

        private void buttonDeleteAll_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("This will delete all Google calendar items created by this program for the Outlook account defined in Settings.\r\n\r\nContinue?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            buttonSyncNow.Enabled = false;
            buttonDeleteAll.Enabled = false;

            textBoxLogs.Clear();
            DateTime SyncStarted = DateTime.Now;

            try
            {
                logboxout("Sync started at " + SyncStarted.ToString());
                logboxout("--------------------------------------------------");

                OutlookCalendar ocal = new OutlookCalendar();

                logboxout("Reading Google Calendar Entries...");
                GoogleCalendar gcal = new GoogleCalendar();
                List<Event> GoogleEntries = gcal.getCalendarEntriesInRange();
                List<Event> GoogleEntriesToDelete = new List<Event>();
                foreach (Event ev in GoogleEntries)
                {
                    if (!string.IsNullOrEmpty(ev.Description) &&
                    ev.Description.Contains("Added by OutlookGoogleSync (" + ocal.AccountName + "):"))
                    {
                        GoogleEntriesToDelete.Add(ev);
                    }
                }

                logboxout("Deleting " + GoogleEntriesToDelete.Count + " Google Calendar Sync Entries...");
                foreach (Event ev in GoogleEntriesToDelete)
                {
                    if (!string.IsNullOrEmpty(ev.Description) &&
                    ev.Description.Contains("Added by OutlookGoogleSync (" + ocal.AccountName + "):"))
                    {
                        gcal.deleteCalendarEntry(ev);
                    }
                }

                logboxout("Done.");
                logboxout("--------------------------------------------------");

                DateTime SyncFinished = DateTime.Now;
                TimeSpan Elapsed = SyncFinished - SyncStarted;
                logboxout("Sync finished at " + SyncFinished.ToString());
                logboxout("Time needed: " + Elapsed.Minutes + " min " + Elapsed.Seconds + " s");
            }
            catch (System.Exception ex)
            {
                logboxout("Error Syncing:\r\n" + ex.ToString());
            }

            buttonSyncNow.Enabled = true;
            buttonDeleteAll.Enabled = true;
        }
    }
}
