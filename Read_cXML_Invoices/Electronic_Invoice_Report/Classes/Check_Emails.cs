using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electronic_Invoice_Report.Objects;
using Utilities;

namespace Electronic_Invoice_Report.Classes
{
    public static class Check_Emails
    {
        private static List<Invoice> invoices = null;
        public static List<Invoice> Check(DateTime date)
        {
            invoices = new List<Invoice>();

            EmailConfig emailConfig = Database.GetEmailConfiguration();
            List<EmailFolderConfig> emailFolders = Database.GetEmailFolderConfigs();
            ReadEmails(emailConfig, emailFolders, date);

            return invoices;
        }

        private static void ReadEmails(EmailConfig config, List<EmailFolderConfig> folders, DateTime date)
        {
            foreach(EmailFolderConfig folder in folders)
            {
                ReadEmail_Report read = new ReadEmail_Report(config.Host, config.Username, config.Password, config.Domain, folder.EmailFolder, 200, date);
                if (read != null)
                {
                    Console.WriteLine(folder.EmailFolder + ": " + read.emailMsgs.Count);
                    for (int i = 0; i < read.emailMsgs.Count; i++)
                    {
                        invoices.Add(new Invoice(
                            folder.EmailFolder.Replace("Invoices", "").Trim(),
                            "EMAIL",
                            read.emailMsgs[i].From.Address,
                            read.emailMsgs[i].Subject,
                            read.emailMsgs[i].Body,
                            read.emailMsgs[i].DateTimeReceived
                            ));
                    }
                }
            }
        }
    }
}
