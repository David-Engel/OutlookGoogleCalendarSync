﻿using System;
using System.Collections.Generic;
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

            //Log on by using a dialog box to choose the profile.
            oNS.Logon("", "", true, true);

            //Alternate logon method that uses a specific profile.
            // If you use this logon method, 
            // change the profile name to an appropriate value.
            //oNS.Logon("YourValidProfile", Missing.Value, false, true); 

            // Get the Calendar folder.
            UseOutlookCalendar = oNS.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            _accountName = oNS.DefaultStore.DisplayName;

            // Done. Log off.
            oNS.Logoff();
        }

        public List<AppointmentItem> getCalendarEntries()
        {
            Items OutlookItems = UseOutlookCalendar.Items;
            if (OutlookItems != null)
            {
                List<AppointmentItem> result = new List<AppointmentItem>();
                foreach (AppointmentItem ai in OutlookItems)
                {
                    result.Add(ai);
                }

                return result;
            }

            return null;
        }

        public List<AppointmentItem> getCalendarEntriesInRange(DateTime syncDateTime)
        {
            List<AppointmentItem> result = new List<AppointmentItem>();

            Items OutlookItems = UseOutlookCalendar.Items;
            OutlookItems.Sort("[Start]", Type.Missing);
            OutlookItems.IncludeRecurrences = true;

            if (OutlookItems != null)
            {
                DateTime min = syncDateTime.AddDays(-Settings.Instance.DaysInThePast);
                DateTime max = syncDateTime.AddDays(+Settings.Instance.DaysInTheFuture + 1);

                //trying this instead, also proposed by WolverineFan, thanks!!! 
                string filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";

                foreach (AppointmentItem ai in OutlookItems.Restrict(filter))
                {
                    result.Add(ai);
                }
            }

            return result;
        }
    }
}
