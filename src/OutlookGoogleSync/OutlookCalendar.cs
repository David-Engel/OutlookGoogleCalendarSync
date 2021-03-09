using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;

namespace OutlookGoogleSync
{
    /// <summary>
    /// Description of OutlookCalendar.
    /// </summary>
    public class OutlookCalendar
    {
        private static OutlookCalendar _instance;

        public static OutlookCalendar Instance
        {
            get
            {
                if (_instance == null) _instance = new OutlookCalendar();
                return _instance;
            }
        }

        public MAPIFolder UseOutlookCalendar;

        private string _accountName = string.Empty;
        public string AccountName
        {
            get { return _accountName; }
        }

        public OutlookCalendar()
        {
            // Create the Outlook application.
            Application oApp = new Application();

            // Get the NameSpace and Logon information.
            NameSpace oNS = oApp.GetNamespace("mapi");

            // Get the Calendar folder.
            UseOutlookCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            Store defaultStore = oNS.DefaultStore;
            _accountName = defaultStore.DisplayName;

            Marshal.ReleaseComObject(defaultStore);
            Marshal.ReleaseComObject(oNS);
            Marshal.ReleaseComObject(oApp);
        }

        public List<AppointmentItem> getCalendarEntries()
        {
            Items outlookItems = UseOutlookCalendar.Items;
            if (outlookItems != null)
            {
                List<AppointmentItem> result = new List<AppointmentItem>();
                foreach (AppointmentItem ai in outlookItems)
                {
                    result.Add(ai);
                }

                Marshal.ReleaseComObject(outlookItems);
                return result;
            }

            return null;
        }

        public List<AppointmentItem> getCalendarEntriesInRange(DateTime syncDateTime)
        {
            List<AppointmentItem> result = new List<AppointmentItem>();

            Items outlookItems = UseOutlookCalendar.Items;
            outlookItems.Sort("[Start]", Type.Missing);
            outlookItems.IncludeRecurrences = true;

            if (outlookItems != null)
            {
                DateTime min = syncDateTime.AddDays(-Settings.Instance.DaysInThePast);
                DateTime max = syncDateTime.AddDays(+Settings.Instance.DaysInTheFuture + 1);

                //trying this instead, also proposed by WolverineFan, thanks!!! 
                string filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";

                foreach (AppointmentItem ai in outlookItems.Restrict(filter))
                {
                    result.Add(ai);
                }

                Marshal.ReleaseComObject(outlookItems);
            }

            return result;
        }
    }
}
