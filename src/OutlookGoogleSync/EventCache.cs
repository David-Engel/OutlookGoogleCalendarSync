using Google.Apis.Calendar.v3.Data;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OutlookGoogleSync
{
    public class EventCache
    {
        private Dictionary<Event, EventCacheEntry> _cache = new Dictionary<Event, EventCacheEntry>();

        public EventCache()
        {
        }

        public EventCacheEntry GetEventCacheEntry(Event @event, string fromAccount)
        {
            if (!_cache.ContainsKey(@event))
            {
                _cache.Add(@event, new EventCacheEntry(@event, fromAccount));
            }

            return _cache[@event];
        }

        public void Clear()
        {
            _cache.Clear();
        }
    }
}
