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
        public static MainForm Instance;

        public const string FILENAME = "settings.xml";

        public Timer ogstimer;
        public DateTime oldtime;
        public List<int> MinuteOffsets = new List<int>();

        private Dictionary<Event, string> _googleEventSignatures = new Dictionary<Event, string>();
        private Dictionary<AppointmentItem, string> _outlookAppointmentSignatures = new Dictionary<AppointmentItem, string>();

        public MainForm()
        {
            InitializeComponent();
            label4.Text = label4.Text.Replace("{version}", System.Windows.Forms.Application.ProductVersion);

            Instance = this;

            //load settings/create settings file
            if (File.Exists(FILENAME))
            {
                Settings.Instance = XMLManager.import<Settings>(FILENAME);
            }
            else
            {
                XMLManager.export(Settings.Instance, FILENAME);
            }

            //update GUI from Settings
            tbDaysInThePast.Text = Settings.Instance.DaysInThePast.ToString();
            tbDaysInTheFuture.Text = Settings.Instance.DaysInTheFuture.ToString();
            tbMinuteOffsets.Text = Settings.Instance.MinuteOffsets;
            cbCalendars.Items.Add(Settings.Instance.SelectedGoogleCalendar);
            cbCalendars.SelectedIndex = 0;
            cbSyncEveryHour.Checked = Settings.Instance.SyncEveryHour;
            cbShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
            cbStartInTray.Checked = Settings.Instance.StartInTray;
            cbMinimizeToTray.Checked = Settings.Instance.MinimizeToTray;
            cbAddDescription.Checked = Settings.Instance.AddDescription;
            cbAddAttendees.Checked = Settings.Instance.AddAttendeesToDescription;
            cbAddReminders.Checked = Settings.Instance.AddReminders;
            cbCreateFiles.Checked = Settings.Instance.CreateTextFiles;

            //Start in tray?
            if (cbStartInTray.Checked)
            {
                this.WindowState = FormWindowState.Minimized;
                notifyIcon1.Visible = true;
                this.Hide();
            }

            //set up timer (every 30s) for checking the minute offsets
            ogstimer = new Timer();
            ogstimer.Interval = 30000;
            ogstimer.Tick += new EventHandler(ogstimer_Tick);
            ogstimer.Start();
            oldtime = DateTime.Now;

            //set up tooltips for some controls
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 10000;
            toolTip1.InitialDelay = 500;
            toolTip1.ReshowDelay = 200;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(cbCalendars,
                "The Google Calendar to synchonize with.");
            toolTip1.SetToolTip(tbMinuteOffsets,
                "One ore more Minute Offsets at which the sync is automatically started each hour. \n" +
                "Separate by comma (e.g. 5,15,25).");
            toolTip1.SetToolTip(cbAddAttendees,
                "While Outlook has fields for Organizer, RequiredAttendees and OptionalAttendees, Google has not.\n" +
                "If checked, this data is added at the end of the description as text.");
            toolTip1.SetToolTip(cbAddReminders,
                "If checked, the reminder set in outlook will be carried over to the Google Calendar entry (as a popup reminder).");
            toolTip1.SetToolTip(cbCreateFiles,
                "If checked, all entries found in Outlook/Google and identified for creation/deletion will be exported \n" +
                "to 4 separate text files in the application's directory (named \"export_*.txt\"). \n" +
                "Only for debug/diagnostic purposes.");
            toolTip1.SetToolTip(cbAddDescription,
                "The description may contain email addresses, which Outlook may complain about (PopUp-Message: \"Allow Access?\" etc.). \n" +
                "Turning this off allows OutlookGoogleSync to run without intervention in this case.");
            toolTip1.SetToolTip(buttonDeleteAll,
                "Delete all calendar items from your Google calendar which were synced from your Outlook calendar in the set time range.");
        }

        void ogstimer_Tick(object sender, EventArgs e)
        {
            if (!cbSyncEveryHour.Checked)
                return;
            DateTime newtime = DateTime.Now;
            if (newtime.Minute != oldtime.Minute)
            {
                oldtime = newtime;
                if (MinuteOffsets.Contains(newtime.Minute))
                {
                    if (cbShowBubbleTooltips.Checked)
                        notifyIcon1.ShowBalloonTip(
                            500,
                            "OutlookGoogleSync",
                            "Sync started at desired minute offset " + newtime.Minute.ToString(),
                            ToolTipIcon.Info
                            );
                    SyncNow_Click(null, null);
                }
            }
        }

        void GetMyGoogleCalendars_Click(object sender, EventArgs e)
        {
            bGetMyCalendars.Enabled = false;
            cbCalendars.Enabled = false;

            try
            {
                GoogleCalendar gcal = new GoogleCalendar();
                List<GoogleCalendarListEntry> calendars = gcal.getCalendars();
                if (calendars != null)
                {
                    cbCalendars.Items.Clear();
                    foreach (GoogleCalendarListEntry mcle in calendars)
                    {
                        cbCalendars.Items.Add(mcle);
                    }
                    MainForm.Instance.cbCalendars.SelectedIndex = 0;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error:\r\n" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            bGetMyCalendars.Enabled = true;
            cbCalendars.Enabled = true;
        }

        void SyncNow_Click(object sender, EventArgs e)
        {
            _googleEventSignatures.Clear();
            _outlookAppointmentSignatures.Clear();

            if (Settings.Instance.SelectedGoogleCalendar.Id == "")
            {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return;
            }

            bSyncNow.Enabled = false;
            buttonDeleteAll.Enabled = false;

            LogBox.Clear();

            DateTime SyncStarted = DateTime.Now;

            try
            {
                logboxout("Sync started at " + SyncStarted.ToString());
                logboxout("--------------------------------------------------");

                logboxout("Reading Outlook Calendar Entries...");
                OutlookCalendar ocal = new OutlookCalendar();
                List<AppointmentItem> OutlookEntries = ocal.getCalendarEntriesInRange();
                if (cbCreateFiles.Checked)
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
                if (cbCreateFiles.Checked)
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
                List<Event> GoogleEntriesToBeDeleted = IdentifyGoogleEntriesToBeDeleted(OutlookEntries, GoogleEntries, accountName);
                if (cbCreateFiles.Checked)
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
                List<AppointmentItem> OutlookEntriesToBeCreated = IdentifyOutlookEntriesToBeCreated(OutlookEntries, GoogleEntries);
                if (cbCreateFiles.Checked)
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
                        if (cbAddDescription.Checked)
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
                        if (cbAddReminders.Checked && ai.ReminderSet)
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

                        if (cbAddAttendees.Checked)
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

            bSyncNow.Enabled = true;
            buttonDeleteAll.Enabled = true;
        }

        //one attendee per line
        public string splitAttendees(string attendees)
        {
            if (attendees == null)
                return "";
            string[] tmp1 = attendees.Split(';');
            for (int i = 0; i < tmp1.Length; i++)
                tmp1[i] = tmp1[i].Trim();
            return String.Join(Environment.NewLine, tmp1);
        }

        public List<Event> IdentifyGoogleEntriesToBeDeleted(List<AppointmentItem> outlook, List<Event> google, string accountName)
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

        public List<AppointmentItem> IdentifyOutlookEntriesToBeCreated(List<AppointmentItem> outlook, List<Event> google)
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
        public string signature(AppointmentItem ai)
        {
            if (!_outlookAppointmentSignatures.ContainsKey(ai))
            {
                _outlookAppointmentSignatures.Add(ai, (GoogleCalendar.GoogleTimeFrom(ai.Start) + ";" + GoogleCalendar.GoogleTimeFrom(ai.End) + ";" + ai.Subject + ";" + ai.Location).Trim());
            }

            return _outlookAppointmentSignatures[ai];
        }

        public string signature(Event ev)
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

        void logboxout(string s)
        {
            LogBox.Text += s + Environment.NewLine;
        }

        void Save_Click(object sender, EventArgs e)
        {
            try
            {
                XMLManager.export(Settings.Instance, FILENAME);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error Saving:\r\n" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void ComboBox1SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Instance.SelectedGoogleCalendar = (GoogleCalendarListEntry)cbCalendars.SelectedItem;
        }

        void TbDaysInThePastTextChanged(object sender, EventArgs e)
        {
            Settings.Instance.DaysInThePast = int.Parse(tbDaysInThePast.Text);
        }

        void TbDaysInTheFutureTextChanged(object sender, EventArgs e)
        {
            Settings.Instance.DaysInTheFuture = int.Parse(tbDaysInTheFuture.Text);
        }

        void TbMinuteOffsetsTextChanged(object sender, EventArgs e)
        {
            Settings.Instance.MinuteOffsets = tbMinuteOffsets.Text;

            MinuteOffsets.Clear();
            char[] delimiters = { ' ', ',', '.', ':', ';' };
            string[] chunks = tbMinuteOffsets.Text.Split(delimiters);
            foreach (string c in chunks)
            {
                int min = 0;
                int.TryParse(c, out min);
                MinuteOffsets.Add(min);
            }
        }


        void CbSyncEveryHourCheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.SyncEveryHour = cbSyncEveryHour.Checked;
        }

        void CbShowBubbleTooltipsCheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.ShowBubbleTooltipWhenSyncing = cbShowBubbleTooltips.Checked;
        }

        void CbStartInTrayCheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.StartInTray = cbStartInTray.Checked;
        }

        void CbMinimizeToTrayCheckedChanged(object sender, System.EventArgs e)
        {
            Settings.Instance.MinimizeToTray = cbMinimizeToTray.Checked;
        }

        void CbAddDescriptionCheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.AddDescription = cbAddDescription.Checked;
        }

        void CbAddRemindersCheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.AddReminders = cbAddReminders.Checked;
        }

        void cbAddAttendees_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.AddAttendeesToDescription = cbAddAttendees.Checked;
        }

        void cbCreateFiles_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Instance.CreateTextFiles = cbCreateFiles.Checked;
        }

        void NotifyIcon1Click(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
        }

        void MainFormResize(object sender, EventArgs e)
        {
            if (!cbMinimizeToTray.Checked)
                return;
            if (this.WindowState == FormWindowState.Minimized)
            {
                notifyIcon1.Visible = true;
                //notifyIcon1.ShowBalloonTip(500, "OutlookGoogleSync", "Click to open again.", ToolTipIcon.Info);
                this.Hide();
            }
            else if (this.WindowState == FormWindowState.Normal)
            {
                notifyIcon1.Visible = false;
            }
        }

        void LinkLabel1LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(linkLabel1.Text);
        }

        private void buttonDeleteAll_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("This will delete all Google calendar items created by this program for the Outlook account defined in Settings.\r\n\r\nContinue?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            bSyncNow.Enabled = false;
            buttonDeleteAll.Enabled = false;

            LogBox.Clear();
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

            bSyncNow.Enabled = true;
            buttonDeleteAll.Enabled = true;
        }
    }
}
