using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes.Report
{
    public class Check_Emails
    {
        private List<Invoice> invoices = null;
        private DateTime AppStarted;
        public List<Invoice> Check(DateTime date, DateTime appStarted)
        {
            AppStarted = appStarted;
            invoices = new List<Invoice>();

            EmailConfig emailConfig = Database.GetEmailConfiguration();
            List<EmailFolderConfig> emailFolders = Database.GetEmailFolderConfigs();
            ReadEmails(emailConfig, emailFolders, date);

            return invoices;
        }

        private void ReadEmails(EmailConfig config, List<EmailFolderConfig> folders, DateTime date)
        {
            foreach(EmailFolderConfig folder in folders)
            {
                ReadEmail_Report read = new ReadEmail_Report(config.Host, config.Username, config.Password, config.Domain, folder.EmailFolder, 200, date);
                if (read != null && read.emailMsgs != null)
                {
                    Console.WriteLine(folder.EmailFolder + ": " + read.emailMsgs.Count);
                    for (int i = 0; i < read.emailMsgs.Count; i++)
                    {
                        if (read.emailMsgs[i].DateTimeReceived < AppStarted)
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
