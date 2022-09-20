namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of Settings.
    /// </summary>
    public class Settings
    {
        private static Settings _instance;

        public static Settings Instance
        {
            get
            {
                if (_instance == null) _instance = new Settings();
                return _instance;
            }
            set
            {
                _instance = value;
            }
        }

        public string MinuteOffsets = "";
        public int DaysInThePast = 1;
        public int DaysInTheFuture = 60;
        public GoogleCalendarListEntry SelectedGoogleCalendar = new GoogleCalendarListEntry();

        public bool SyncEveryHour = false;
        public bool ShowBubbleTooltipWhenSyncing = false;
        public bool StartInTray = false;
        public bool MinimizeToTray = false;

        public bool AddDescription = true;
        public bool AddReminders = false;
        public bool AddAttendeesToDescription = true;
        public bool CreateTextFiles = true;

        public Settings()
        {
        }
    }
}
