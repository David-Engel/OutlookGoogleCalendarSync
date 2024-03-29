﻿using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class MainForm : Form
    {
        internal const string FILENAME = "settings.xml";

        private readonly Timer _ogstimer;
        private DateTime _oldtime;
        private readonly List<int> _minuteOffsets = new List<int>();
        private FormWindowState _previousWindowState = FormWindowState.Normal;
        private readonly BackgroundWorker _syncWorker = new BackgroundWorker();

        public MainForm()
        {
            InitializeComponent();
            labelAbout.Text = labelAbout.Text.Replace("{version}", System.Windows.Forms.Application.ProductVersion);

            //load settings/create settings file
            if (File.Exists(FILENAME))
            {
                Settings.Instance = XMLManager.Import<Settings>(FILENAME);
            }
            else
            {
                XMLManager.Export(Settings.Instance, FILENAME);
            }

            //update GUI from Settings
            numericUpDownDaysInThePast.Value = Settings.Instance.DaysInThePast;
            numericUpDownDaysInTheFuture.Value = Settings.Instance.DaysInTheFuture;
            textBoxMinuteOffsets.Text = Settings.Instance.MinuteOffsets;
            comboBoxGoogleCalendars.Items.Add(Settings.Instance.SelectedGoogleCalendar);
            comboBoxGoogleCalendars.SelectedIndex = 0;
            checkBoxSyncEveryHour.Checked = Settings.Instance.SyncEveryHour;
            checkBoxShowBubbleTooltips.Checked = Settings.Instance.ShowBubbleTooltipWhenSyncing;
            checkBoxStartInTray.Checked = Settings.Instance.StartInTray;
            checkBoxMinimizeToTray.Checked = Settings.Instance.MinimizeToTray;
            checkBoxAddDescription.Checked = Settings.Instance.AddDescription;
            checkBoxAddAttendees.Checked = Settings.Instance.AddAttendeesToDescription;
            checkBoxAddReminders.Checked = Settings.Instance.AddReminders;
            checkBoxCreateFiles.Checked = Settings.Instance.CreateTextFiles;

            comboBoxOutlookCalendars.Items.Clear();
            comboBoxOutlookCalendars.Items.Add(Settings.Instance.OutlookCalendarToSync);
            comboBoxOutlookCalendars.SelectedIndex = 0;

            //set up timer (every 30s) for checking the minute offsets
            _ogstimer = new Timer();
            _ogstimer.Interval = 30000;
            _ogstimer.Tick += new EventHandler(Ogstimer_Tick);
            _ogstimer.Start();
            _oldtime = DateTime.Now;

            _syncWorker.WorkerReportsProgress = true;
            _syncWorker.WorkerSupportsCancellation = true;
            _syncWorker.DoWork += SyncWorker_DoWork;
            _syncWorker.ProgressChanged += syncWorker_ProgressChanged;
            _syncWorker.RunWorkerCompleted += syncWorker_RunWorkerCompleted;

            //set up tooltips for some controls
            ToolTip toolTip1 = new ToolTip
            {
                AutoPopDelay = 10000,
                InitialDelay = 500,
                ReshowDelay = 200,
                ShowAlways = true
            };
            toolTip1.SetToolTip(comboBoxGoogleCalendars,
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

        private void MainForm_Load(object sender, EventArgs e)
        {
            //Start in tray?
            if (checkBoxStartInTray.Checked)
            {
                WindowState = FormWindowState.Minimized;
            }
        }

        private void Ogstimer_Tick(object sender, EventArgs e)
        {
            if (!checkBoxSyncEveryHour.Checked)
            {
                return;
            }

            DateTime newtime = DateTime.Now;
            if (newtime.Minute != _oldtime.Minute)
            {
                _oldtime = newtime;
                if (_minuteOffsets.Contains(newtime.Minute))
                {
                    if (checkBoxShowBubbleTooltips.Checked)
                    {
                        notifyIcon1.ShowBalloonTip(
                            500,
                            "OutlookGoogleSync",
                            "Sync started at desired minute offset " + newtime.Minute.ToString(),
                            ToolTipIcon.Info
                            );
                    }

                    ButtonSyncNow_Click(null, null);
                }
            }
        }

        private void ButtonGetMyGoogleCalendars_Click(object sender, EventArgs e)
        {
            buttonGetMyCalendars.Enabled = false;
            comboBoxGoogleCalendars.Enabled = false;

            try
            {
                GoogleCalendar gcal = new GoogleCalendar();
                List<GoogleCalendarListEntry> calendars = gcal.GetCalendars();
                if (calendars != null)
                {
                    comboBoxGoogleCalendars.Items.Clear();
                    foreach (GoogleCalendarListEntry mcle in calendars)
                    {
                        comboBoxGoogleCalendars.Items.Add(mcle);
                    }

                    comboBoxGoogleCalendars.SelectedIndex = 0;
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("Error:\r\n" + ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            buttonGetMyCalendars.Enabled = true;
            comboBoxGoogleCalendars.Enabled = true;
        }

        private void ButtonSyncNow_Click(object sender, EventArgs e)
        {
            if (_syncWorker.IsBusy)
            {
                return;
            }

            if (Settings.Instance.SelectedGoogleCalendar.Id == "")
            {
                MessageBox.Show("You need to select a Google Calendar first on the 'Settings' tab.");
                return;
            }

            buttonSyncNow.Enabled = false;
            buttonDeleteAll.Enabled = false;

            textBoxLogs.Clear();

            _syncWorker.RunWorkerAsync();
        }

        private void SyncWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            if (e.Argument is string @string &&
                "DELETE".Equals(@string))
            {
                deleteAllSyncItems();
                return;
            }

            AppointmentItemCache appointmentItemCache = new AppointmentItemCache();
            EventCache eventCache = new EventCache();

            DateTime syncStarted = DateTime.Now;
            OutlookCalendar ocal = Settings.Instance.OutlookCalendarToSync;

            try
            {
                logboxout("Sync started at " + syncStarted.ToString());
                logboxout("--------------------------------------------------");

                logboxout("Reading Outlook Calendar Entries...");
                List<AppointmentItemCacheEntry> OutlookEntries = new List<AppointmentItemCacheEntry>();
                foreach (AppointmentItem a in Settings.Instance.OutlookCalendarToSync.GetAppointmentItemsInRange(syncStarted))
                {
                    OutlookEntries.Add(appointmentItemCache.GetAppointmentItemCacheEntry(a, ocal.ToString()));
                }

                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_found_in_outlook.txt"))
                    {
                        foreach (AppointmentItemCacheEntry ai in OutlookEntries)
                        {
                            tw.WriteLine(ai.Signature);
                        }
                    }
                }

                logboxout("Found " + OutlookEntries.Count + " Outlook Calendar Entries.");
                logboxout("--------------------------------------------------");
                logboxout("Reading Google Calendar Entries...");

                GoogleCalendar gcal = new GoogleCalendar();
                List<EventCacheEntry> GoogleEntries = new List<EventCacheEntry>();
                foreach (Event ev in gcal.GetCalendarEntriesInRange(syncStarted))
                {
                    GoogleEntries.Add(eventCache.GetEventCacheEntry(ev, ocal.ToString()));
                }

                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_found_in_google.txt"))
                    {
                        foreach (EventCacheEntry ev in GoogleEntries)
                        {
                            tw.WriteLine(ev.Signature);
                        }
                    }
                }

                logboxout("Found " + GoogleEntries.Count + " Google Calendar Entries.");
                logboxout("--------------------------------------------------");

                List<EventCacheEntry> GoogleEntriesToBeDeleted = identifyGoogleEntriesToBeDeleted(OutlookEntries, GoogleEntries);
                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_to_be_deleted.txt"))
                    {
                        foreach (EventCacheEntry ev in GoogleEntriesToBeDeleted)
                        {
                            tw.WriteLine(ev.Signature);
                        }
                    }
                }

                logboxout(GoogleEntriesToBeDeleted.Count + " Google Calendar Entries to be deleted.");

                //OutlookEntriesToBeCreated ...in Google!
                List<AppointmentItemCacheEntry> OutlookEntriesToBeCreated = identifyOutlookEntriesToBeCreated(OutlookEntries, GoogleEntries);
                if (checkBoxCreateFiles.Checked)
                {
                    using (TextWriter tw = new StreamWriter("export_to_be_created.txt"))
                    {
                        foreach (AppointmentItemCacheEntry ai in OutlookEntriesToBeCreated)
                        {
                            tw.WriteLine(ai.Signature);
                        }
                    }
                }

                logboxout(OutlookEntriesToBeCreated.Count + " Entries to be created in Google.");
                logboxout("--------------------------------------------------");

                if (GoogleEntriesToBeDeleted.Count > 0)
                {
                    logboxout("Deleting " + GoogleEntriesToBeDeleted.Count + " Google Calendar Entries...");
                    foreach (EventCacheEntry ev in GoogleEntriesToBeDeleted)
                        gcal.DeleteCalendarEntry(ev.Event);
                    logboxout("Done.");
                    logboxout("--------------------------------------------------");
                }

                if (OutlookEntriesToBeCreated.Count > 0)
                {
                    logboxout("Creating " + OutlookEntriesToBeCreated.Count + " Entries in Google...");
                    foreach (AppointmentItemCacheEntry aice in OutlookEntriesToBeCreated)
                    {
                        AppointmentItem ai = aice.AppointmentItem;
                        Event ev = new Event
                        {
                            Start = new EventDateTime(),
                            End = new EventDateTime()
                        };

                        if (ai.AllDayEvent)
                        {
                            ev.Start.Date = ai.Start.ToString("yyyy-MM-dd");
                            ev.End.Date = ai.End.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            ev.Start.DateTime = ai.Start;
                            ev.End.DateTime = ai.End;
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
                            EventReminder reminder = new EventReminder();
                            reminder.Method = "popup";
                            reminder.Minutes = ai.ReminderMinutesBeforeStart;
                            ev.Reminders = new Event.RemindersData();
                            ev.Reminders.UseDefault = false;
                            ev.Reminders.Overrides = new List<EventReminder>();
                            ev.Reminders.Overrides.Add(reminder);
                        }
                        else
                        {
                            ev.Reminders = new Event.RemindersData();
                            ev.Reminders.UseDefault = false;
                        }

                        ev.Description += Environment.NewLine;
                        ev.Description += Environment.NewLine + "==============================================";
                        ev.Description += Environment.NewLine + "Added by OutlookGoogleSync (" + ocal.ToString() + "):" + Environment.NewLine;

                        if (checkBoxAddAttendees.Checked)
                        {
                            ev.Description += Environment.NewLine + "ORGANIZER: " + Environment.NewLine + ai.Organizer + Environment.NewLine;
                            ev.Description += Environment.NewLine + "REQUIRED: " + Environment.NewLine + splitAttendees(ai.RequiredAttendees) + Environment.NewLine;
                            if (ai.OptionalAttendees != null)
                            {
                                ev.Description += Environment.NewLine + "OPTIONAL: " + Environment.NewLine + splitAttendees(ai.OptionalAttendees);
                            }
                        }

                        ev.Description += Environment.NewLine + "==============================================";

                        gcal.AddEntry(ev);
                    }

                    logboxout("Done.");
                    logboxout("--------------------------------------------------");
                }

                DateTime syncFinished = DateTime.Now;
                TimeSpan elapsed = syncFinished - syncStarted;
                logboxout("Sync finished at " + syncFinished.ToString());
                logboxout("Time needed: " + elapsed.Minutes + " min " + elapsed.Seconds + " s");
            }
            catch (System.Exception ex)
            {
                logboxout("Error Syncing:\r\n" + ex.ToString());
            }

            eventCache.Clear();
            freeCOMResources(ocal, appointmentItemCache);
        }

        private void freeCOMResources(OutlookCalendar oc, AppointmentItemCache appointmentItemCache)
        {
            try
            {
                appointmentItemCache?.ClearAndReleaseAll();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (System.Exception ex)
            {
                logboxout("Warning: Error freeing COM resources:\r\n" + ex.ToString());
            }
        }

        //one attendee per line
        private string splitAttendees(string attendees)
        {
            if (attendees == null)
            {
                return "";
            }

            string[] tmp1 = attendees.Split(';');
            for (int i = 0; i < tmp1.Length; i++)
            {
                tmp1[i] = tmp1[i].Trim();
            }

            return String.Join(Environment.NewLine, tmp1);
        }

        private List<EventCacheEntry> identifyGoogleEntriesToBeDeleted(List<AppointmentItemCacheEntry> outlook, List<EventCacheEntry> google)
        {
            List<EventCacheEntry> result = new List<EventCacheEntry>();
            foreach (EventCacheEntry g in google)
            {
                bool found = false;
                foreach (AppointmentItemCacheEntry o in outlook)
                {
                    if (g.Signature == o.Signature)
                    {
                        found = true;
                    }
                }

                if (!found && g.IsSyncItem)
                {
                    result.Add(g);
                }
            }

            return result;
        }

        private List<AppointmentItemCacheEntry> identifyOutlookEntriesToBeCreated(List<AppointmentItemCacheEntry> outlook, List<EventCacheEntry> google)
        {
            List<AppointmentItemCacheEntry> result = new List<AppointmentItemCacheEntry>();
            foreach (AppointmentItemCacheEntry o in outlook)
            {
                bool found = false;
                foreach (EventCacheEntry g in google)
                {
                    if (g.Signature == o.Signature)
                    {
                        found = true;
                    }
                }

                if (!found)
                {
                    result.Add(o);
                }
            }

            return result;
        }

        private void logboxout(string s)
        {
            _syncWorker.ReportProgress(0, s);
        }

        private void syncWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string s = (string)e.UserState;
            textBoxLogs.Text += s + Environment.NewLine;
            textBoxLogs.SelectionStart = textBoxLogs.Text.Length - 1;
            textBoxLogs.SelectionLength = 0;
            textBoxLogs.ScrollToCaret();
        }

        private void syncWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            buttonDeleteAll.Enabled = true;
            buttonSyncNow.Enabled = true;
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
            Settings.Instance.SelectedGoogleCalendar = (GoogleCalendarListEntry)comboBoxGoogleCalendars.SelectedItem;
        }

        private void numericUpDownDaysInThePast_ValueChanged(object sender, EventArgs e)
        {
            Settings.Instance.DaysInThePast = (int)numericUpDownDaysInThePast.Value;
        }

        private void numericUpDownDaysInTheFuture_ValueChanged(object sender, EventArgs e)
        {
            Settings.Instance.DaysInTheFuture = (int)numericUpDownDaysInTheFuture.Value;
        }

        private void textBoxMinuteOffsets_TextChanged(object sender, EventArgs e)
        {
            Settings.Instance.MinuteOffsets = textBoxMinuteOffsets.Text;

            _minuteOffsets.Clear();
            char[] delimiters = { ' ', ',', '.', ':', ';' };
            string[] chunks = textBoxMinuteOffsets.Text.Split(delimiters);
            foreach (string c in chunks)
            {
                int.TryParse(c, out int min);
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
            if (_syncWorker.IsBusy)
            {
                return;
            }

            if (MessageBox.Show("This will delete all Google calendar items created by this program for the Outlook account defined in Settings.\r\n\r\nContinue?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) != System.Windows.Forms.DialogResult.Yes)
            {
                return;
            }

            buttonSyncNow.Enabled = false;
            buttonDeleteAll.Enabled = false;

            textBoxLogs.Clear();
            _syncWorker.RunWorkerAsync("DELETE");
        }

        private void deleteAllSyncItems()
        {
            DateTime syncStarted = DateTime.Now;
            OutlookCalendar ocal = Settings.Instance.OutlookCalendarToSync;

            try
            {
                logboxout("Sync started at " + syncStarted.ToString());
                logboxout("--------------------------------------------------");

                logboxout("Reading Google Calendar Entries...");
                GoogleCalendar gcal = new GoogleCalendar();
                List<Event> GoogleEntries = gcal.GetCalendarEntriesInRange(syncStarted);
                List<Event> GoogleEntriesToDelete = new List<Event>();
                foreach (Event ev in GoogleEntries)
                {
                    if (!string.IsNullOrEmpty(ev.Description) &&
                    ev.Description.Contains("Added by OutlookGoogleSync ("))
                    {
                        GoogleEntriesToDelete.Add(ev);
                    }
                }

                logboxout("Deleting " + GoogleEntriesToDelete.Count + " Google Calendar Sync Entries...");
                foreach (Event ev in GoogleEntriesToDelete)
                {
                    gcal.DeleteCalendarEntry(ev);
                }

                logboxout("Done.");
                logboxout("--------------------------------------------------");

                DateTime syncFinished = DateTime.Now;
                TimeSpan elapsed = syncFinished - syncStarted;
                logboxout("Sync finished at " + syncFinished.ToString());
                logboxout("Time needed: " + elapsed.Minutes + " min " + elapsed.Seconds + " s");
            }
            catch (System.Exception ex)
            {
                logboxout("Error Syncing:\r\n" + ex.ToString());
            }

            freeCOMResources(ocal, null);
        }

        private void comboBoxOutlookCalendars_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Instance.OutlookCalendarToSync = (OutlookCalendar)comboBoxOutlookCalendars.SelectedItem;
        }

        private void buttonLoadOutlookCalendars_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            OutlookCalendar saved = Settings.Instance.OutlookCalendarToSync;
            try
            {
                comboBoxOutlookCalendars.Items.Clear();
                foreach (OutlookCalendar cal in OutlookHelper.GetCalendars())
                {
                    int position = comboBoxOutlookCalendars.Items.Add(cal);
                    if (cal.ToString().Equals(saved.ToString()))
                    {
                        comboBoxOutlookCalendars.SelectedIndex = position;
                    }
                }
                Cursor.Current = Cursors.Default;
            }
            catch (System.Exception ex)
            {
                Cursor.Current = Cursors.Default;
                MessageBox.Show(ex.ToString());
            }
        }
    }
}
