using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;

namespace authAccess
{
    public class OutlookEmails
    {
        private static string sFilter;
        public string EmailFrom { get; set; }
        public string EmailSubject { get; set; }
        public string EmailBody { get; set; }
        public string EmailTo { get; set; }
        public DateTime EmailDate { get; set; }

        public static List<OutlookEmails> ReadMailItems()
        { 
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;

            Items mailItems = null;
            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails  emailDetails;
            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                //mailItems = inboxFolder.Items;
                 DateTime startTime = DateTime.Now;
                 DateTime endTime = startTime.AddDays(-150);
                 /*sFilter = "[ReceivedTime] >= '" 
                                + startTime.ToString("MM/dd/yyyy HH:mm")
                                + "' AND [ReceivedTime] <= '"
                                + endTime.ToString("MM/dd/yyyy HH:mm") + "'";
                 */
                sFilter = "[ReceivedTime] <= '"
                                + startTime.ToString("MM/dd/yyyy HH:mm")
                                + "' AND [ReceivedTime] >= '"
                                + endTime.ToString("MM/dd/yyyy HH:mm") + "'";
                //    04/30/2020 14:53
                //sFilter = "[ReceivedTime] >= '07/05/2020 00:00' AND [ReceivedTime] <= '09/05/2020 00:00' ";

                //mailItems = mailItems.Restrict("[ReceivedTime] > '" + dt.ToString("MM/dd/yyyy hh:mm:ss tt") + "'");
                //mailItems = mailItems.Restrict("[ReceivedTime] > '" + dt.ToString("MM/dd/yyyy HH:mm") + "'");
                mailItems = inboxFolder.Items.Restrict(sFilter);
                 //mailItems.Restrict(sFilter);
                 //foreach (dynamic item in mailItems)
                foreach (dynamic item in mailItems)
                  {
                    if (item is MailItem)
                    {
                        emailDetails = new OutlookEmails();
                        emailDetails.EmailFrom = item.SenderEmailAddress;
                        emailDetails.EmailSubject = item.Subject;
                        emailDetails.EmailBody = item.Body;
                        emailDetails.EmailTo = item.To;
                        emailDetails.EmailDate = item.ReceivedTime;

                        listEmailDetails.Add(emailDetails);

                        ReleaseComObject(item);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                ReleaseComObject(mailItems);
                ReleaseComObject(inboxFolder);
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
            return listEmailDetails;
        }
        private static void ReleaseComObject(object obj)
        {
            if(obj !=null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
