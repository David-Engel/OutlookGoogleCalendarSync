using Google.Apis.Calendar.v3.Data;
using System;

namespace OutlookGoogleSync
{
    internal class EventCacheEntry
    {
        internal EventCacheEntry(Event @event, string fromAccount)
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
                    Event.Start.DateTime = DateTime.Parse(Event.Start.Date);
                }
                if (Event.End.DateTime == null)
                {
                    Event.End.DateTime = DateTime.Parse(Event.End.Date);
                }

                Signature = (Event.Start.DateTime + ";" + Event.End.DateTime + ";" + Event.Summary + ";" + Event.Location).Trim();
            }
            else
            {
                Signature = (Event.Status + ";" + Event.Summary + ";" + Event.Location).Trim();
            }
        }

        internal Event Event { get; }

        internal bool IsSyncItem { get; } = false;

        internal string FromAccount { get; }

        internal string Signature { get; }
    }
}
