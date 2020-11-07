using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace WarrantyDataMiner
{
    class Program
    {
        static void Main(string[] args)
        {
            var mails = OutlookEmails.ReadMailItems();
            List<string> warrantyBodies = new List<string>();

            int i = 0;

            foreach (var mail in mails)
            {
                if (mail.EmailSubject != null)
                {
                    if (mail.EmailSubject.StartsWith("New Product Registration"))
                    {
                        warrantyBodies.Add(mail.EmailBody);

                        i++;
                    }
                }                
            }

            Console.WriteLine("");
            Console.WriteLine(i + " Warranty Emails Found.");

            Console.WriteLine("");
            Console.WriteLine("Writing Email Content to Excel Records...");

            Application excel = new Application();
            excel.Visible = true;
            Workbook wb = excel.Workbooks.Open("WarrantyRecords");
            Worksheet sh = wb.ActiveSheet;
            excel.WindowState = XlWindowState.xlMaximized;

            foreach (var warrantyBody in warrantyBodies)
            {
                string[] splitBodyByReturn = null;

                if (warrantyBody.Contains("Subject: New Product Registration") || warrantyBody.Contains("Subject: [SPAM] New Product Registration"))
                {
                    splitBodyByReturn = warrantyBody.Substring(warrantyBody.IndexOf("ProductRegistrationID:")).Split(new string[] { "\n" }, StringSplitOptions.None);
                }
                else
                {
                    splitBodyByReturn = warrantyBody.Split(new string[] { "\n" }, StringSplitOptions.None);
                }
                
                List<string> recordEntries = new List<string>();
                List<string> recordEntriesSplitByColon = new List<string>();

                // Create list of each line entry from email body
                for (int j = 0; j < 19; j++)
                {
                    recordEntries.Add(splitBodyByReturn[j * 2]);
                }

                for (int k = 0; k < 19; k++)
                {
                    if (recordEntries[k].Contains("Form inserted: "))
                    {
                        recordEntriesSplitByColon.Add(recordEntries[k].Replace("Form inserted: ", "").Trim());
                    }
                    else if (recordEntries[k].Contains("Form updated: "))
                    {
                        recordEntriesSplitByColon.Add(recordEntries[k].Replace("Form updated: ", "").Trim());
                    }
                    else if (recordEntries[k].Contains("<mailto"))
                    {
                        string[] recordSplitByColon = recordEntries[k].Split(':');
                        if (recordSplitByColon.Length > 0) recordEntriesSplitByColon.Add(recordSplitByColon[1].Replace("<mailto", "").Trim());
                    }
                    else 
                    {
                        string[] recordSplitByColon = recordEntries[k].Split(':');
                        if (recordSplitByColon.Length > 0) recordEntriesSplitByColon.Add(recordSplitByColon[1].Trim());
                    }
                }

                Range line = (Range)sh.Rows[2];
                line.Insert();
                
                sh.Cells[2, "A"].Value2 = recordEntriesSplitByColon[0];
                sh.Cells[2, "B"].Value2 = recordEntriesSplitByColon[1];
                sh.Cells[2, "C"].Value2 = recordEntriesSplitByColon[2];
                sh.Cells[2, "D"].Value2 = recordEntriesSplitByColon[3];
                sh.Cells[2, "E"].Value2 = recordEntriesSplitByColon[4];
                sh.Cells[2, "F"].Value2 = recordEntriesSplitByColon[5];
                sh.Cells[2, "G"].Value2 = recordEntriesSplitByColon[6];
                sh.Cells[2, "H"].Value2 = recordEntriesSplitByColon[7];
                sh.Cells[2, "I"].Value2 = recordEntriesSplitByColon[8];
                sh.Cells[2, "J"].Value2 = recordEntriesSplitByColon[9];
                sh.Cells[2, "K"].Value2 = recordEntriesSplitByColon[10];
                sh.Cells[2, "L"].Value2 = recordEntriesSplitByColon[11];
                sh.Cells[2, "M"].Value2 = recordEntriesSplitByColon[12];
                sh.Cells[2, "N"].Value2 = recordEntriesSplitByColon[13];
                sh.Cells[2, "O"].Value2 = recordEntriesSplitByColon[14];
                sh.Cells[2, "P"].Value2 = recordEntriesSplitByColon[15];
                sh.Cells[2, "Q"].Value2 = recordEntriesSplitByColon[16];
                sh.Cells[2, "R"].Value2 = recordEntriesSplitByColon[17];
                sh.Cells[2, "S"].Value2 = recordEntriesSplitByColon[18];
            }

            //wb.Close(true);
            //excel.Quit();

            Console.WriteLine("");
            Console.WriteLine("Entry of Records Complete.");
            Console.ReadKey();
        }
    }
}