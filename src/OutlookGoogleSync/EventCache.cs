using Google.Apis.Calendar.v3.Data;
using System.Collections.Generic;

namespace OutlookGoogleSync
{
    internal class EventCache
    {
        private Dictionary<Event, EventCacheEntry> _cache = new Dictionary<Event, EventCacheEntry>();

        internal EventCache()
        {
        }

        internal EventCacheEntry GetEventCacheEntry(Event @event, string fromAccount)
        {
            if (!_cache.ContainsKey(@event))
            {
                _cache.Add(@event, new EventCacheEntry(@event, fromAccount));
            }

            return _cache[@event];
        }

        internal void Clear()
        {
            _cache.Clear();
        }
    }
}
