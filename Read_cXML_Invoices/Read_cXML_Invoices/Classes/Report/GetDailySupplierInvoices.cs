using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes.Report
{
    public static class GetDailySupplierInvoices
    {
        public static List<Invoice> GetThem(DateTime appStarted)
        {
            List<Invoice> Invoices = new List<Invoice>();
            DateTime temp = DateTime.Now.AddDays(-5);
            DateTime check = new DateTime(temp.Year, temp.Month, temp.Day, 0, 0, 0);

            try
            {
                Check_cXML ccxml = new Check_cXML();
                Console.WriteLine($"DateTime Started: {DateTime.Now}");
                Invoices.AddRange(ccxml.Check(check, appStarted));
                Console.WriteLine("XML invoices captured");

                Check_EDI cedi = new Check_EDI();
                Invoices.AddRange(cedi.Check(check, appStarted));
                Console.WriteLine("EDI Invoices captured");

                Check_Emails cemail = new Check_Emails();
                Invoices.AddRange(cemail.Check(check, appStarted));
                Console.WriteLine("Email Invoices captured");

                foreach (Invoice invoice in Invoices)
                    Database.InsertInvoice(invoice);
                Console.WriteLine("Invoices inserted to db");
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, "GetDailySupplierInvoices", "GetThem"));
            }

            return Invoices;
        }
    }
}
