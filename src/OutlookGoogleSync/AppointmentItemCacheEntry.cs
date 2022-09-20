using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleSync
{
    internal class AppointmentItemCacheEntry
    {
        internal AppointmentItemCacheEntry(AppointmentItem appointmentItem, string fromAccount)
        {
            AppointmentItem = appointmentItem;
            FromAccount = fromAccount;

            Signature = (appointmentItem.Start + ";" +
                appointmentItem.End + ";" +
                appointmentItem.Subject + ";" + appointmentItem.Location).Trim();
        }

        internal AppointmentItem AppointmentItem { get; }

        internal string FromAccount { get; }

        internal string Signature { get; }
    }
}
