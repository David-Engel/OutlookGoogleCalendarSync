using Google.Apis.Calendar.v3.Data;
using System;

namespace OutlookGoogleSync
{
    public class EventCacheEntry
    {
        public EventCacheEntry(Event @event, string fromAccount)
        {
            Event = @event;
            FromAccount = fromAccount;
            if (!string.IsNullOrEmpty(@event.Description) &&
                    Event.Description.Contains("Added by OutlookGoogleSync (" + fromAccount + "):"))
            {
                IsSyncItem = true;
            }

            if (Event.Start != null)
            {
                if (Event.Start.DateTime == null)
                {
                    Event.Start.DateTime = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(Event.Start.Date));
                }
                if (Event.End.DateTime == null)
                {
                    Event.End.DateTime = GoogleCalendar.GoogleTimeFrom(DateTime.Parse(Event.End.Date));
                }

                Signature = (Event.Start.DateTime + ";" + Event.End.DateTime + ";" + Event.Summary + ";" + Event.Location).Trim();
            }
            else
            {
                Signature = (Event.Status + ";" + Event.Summary + ";" + Event.Location).Trim();
            }
        }

        public Event Event { get; }

        public bool IsSyncItem { get; } = false;

        public string FromAccount { get; }

        public string Signature { get; }
    }
}
