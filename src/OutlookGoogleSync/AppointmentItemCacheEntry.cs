using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleSync
{
    public class AppointmentItemCacheEntry
    {
        public AppointmentItemCacheEntry(AppointmentItem appointmentItem, string fromAccount)
        {
            AppointmentItem = appointmentItem;
            FromAccount = fromAccount;

            Signature = (GoogleCalendar.GoogleTimeFrom(appointmentItem.Start) + ";" +
                GoogleCalendar.GoogleTimeFrom(appointmentItem.End) + ";" +
                appointmentItem.Subject + ";" + appointmentItem.Location).Trim();
        }

        public AppointmentItem AppointmentItem { get; }

        public string FromAccount { get; }

        public string Signature { get; }
    }
}
