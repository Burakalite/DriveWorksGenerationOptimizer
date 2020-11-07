using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WarrantyParser
{
    public class OutlookEmails
    {
        public string EmailSubject { get; set; }

        public string EmailBody { get; set; }

        public static List<OutlookEmails> ReadMailItems()
        {
            Application outlookApplication = null;
            NameSpace outlookNamespace = null;
            MAPIFolder inboxFolder = null;
            Items mailItems = null;

            List<OutlookEmails> listEmailDetails = new List<OutlookEmails>();
            OutlookEmails emailDetails;

            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");

                inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                mailItems = inboxFolder.Items;

                foreach (Object mail in mailItems)
                {
                    if ((mail as MailItem) != null)
                    {
                        emailDetails = new OutlookEmails();
                        emailDetails.EmailSubject = (mail as MailItem).Subject;
                        emailDetails.EmailBody = (mail as MailItem).Body;

                        listEmailDetails.Add(emailDetails);

                        ReleaseComObject(mail);
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
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
        }
    }
}
