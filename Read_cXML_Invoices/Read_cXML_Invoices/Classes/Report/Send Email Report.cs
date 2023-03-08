using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes.Report
{
    public class Send_Email_Report
    {
        private List<InvoiceHeader> invoicesToCheck = new List<InvoiceHeader>();
        public Send_Email_Report(InvoiceHeader[] invoices)
        {
            PopulateInvoices(invoices);
        }

        private void PopulateInvoices(InvoiceHeader[] invoices)
        {
            foreach(InvoiceHeader inv in invoices)
            {
                if (inv.Vendor == "Digi-Key")
                    invoicesToCheck.Add(inv);
            }
        }

        private void GetNumberOfEmails()
        {
            int count = 0;
            try
            {
                EmailConfig config = Database.GetEmailConfiguration();
                if (config != null)
                {
                    ReadEmail readEmail = new ReadEmail(config.Host, config.Username, config.Password, config.Domain, "Downloaded", 0, "");
                    for (int i = 0; i < readEmail.emailMsgs.Count; i++)
                    {
                        if (readEmail.emailMsgs[i].Subject.StartsWith("Digi-Key") && readEmail.emailMsgs[i].DateTimeReceived < DateTime.Now)
                        {
                            count++;
                            readEmail.UpdateEmailStatus(i);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, "Send_Email_Report", "GetNumberOfEmails"));
            }
        }
    }
}
