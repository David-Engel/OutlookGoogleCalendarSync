using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OutlookGoogleSync
{
    [Serializable()]
    public class OutlookCalendar
    {
        public string MailboxName;

        public string FolderName;

        public MAPIFolder GetFolder()
        {
            return OutlookHelper.GetFolder(this);
        }

        public OutlookCalendar()
        {
        }

        public OutlookCalendar(string mailboxName, string folderName)
        {
            FolderName = folderName;
            MailboxName = mailboxName;
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(MailboxName) || string.IsNullOrEmpty(FolderName))
            {
                return "";
            }

            return MailboxName + " - " + FolderName;
        }

        public List<AppointmentItem> GetAppointmentItemsInRange(DateTime syncDateTime)
        {
            List<AppointmentItem> result = new List<AppointmentItem>();

            if (FolderName == null || MailboxName == null)
            {
                return result;
            }

            MAPIFolder folder = GetFolder();
            if (folder == null)
            {
                return result;
            }

            try
            {
                Items outlookItems = folder.Items;
                outlookItems.Sort("[Start]", Type.Missing);
                outlookItems.IncludeRecurrences = true;

                if (outlookItems != null)
                {
                    try
                    {
                        DateTime min = syncDateTime.AddDays(-Settings.Instance.DaysInThePast);
                        DateTime max = syncDateTime.AddDays(+Settings.Instance.DaysInTheFuture + 1);
                        string filter = "[End] >= '" + min.ToString("g") + "' AND [Start] < '" + max.ToString("g") + "'";
                        foreach (AppointmentItem ai in outlookItems.Restrict(filter))
                        {
                            result.Add(ai);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(outlookItems);
                    }
                }
            }
            finally
            {
                if (folder != null)
                {
                    Marshal.ReleaseComObject(folder);
                }
            }

            return result;
        }
    }
}
