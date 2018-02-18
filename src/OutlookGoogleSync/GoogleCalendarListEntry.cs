using Google.Apis.Calendar.v3.Data;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of MyCalendarListEntry.
    /// </summary>
    public class GoogleCalendarListEntry
    {
        public string Id = "";
        public string Name = "";

        public GoogleCalendarListEntry()
        {
        }

        public GoogleCalendarListEntry(CalendarListEntry init)
        {
            Id = init.Id;
            Name = init.Summary;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
