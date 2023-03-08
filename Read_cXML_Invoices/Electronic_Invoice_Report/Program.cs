using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electronic_Invoice_Report.Classes;
using Electronic_Invoice_Report.Objects;
using Utilities;

namespace Electronic_Invoice_Report
{
    class Program
    {
        public List<Invoice> Invoices = new List<Invoice>();
        public Program(DateTime check)
        {
            try
            {
                Console.WriteLine($"DateTime Started: {DateTime.Now}");
                Invoices.AddRange(Check_cXML.Check(check));
                Console.WriteLine("XML invoices captured");

                Invoices.AddRange(Check_Emails.Check(check));
                Console.WriteLine("Email Invoices captured");

                foreach (Invoice invoice in Invoices)
                    Database.InsertInvoice(invoice);
                Console.WriteLine("Invoices inserted to db");

                SendReport.SendIt(Invoices);
                Console.WriteLine("Report sent");
                Console.WriteLine($"DateTime Ended: {DateTime.Now}");
            }
            catch(Exception ex)
            {
                Email.SendErrorMessage(ex, "Electronic_Invoice_Report", "Program.Program", null);
            }
        }
        public static void Main(string[] args)
        {
            DateTime temp = DateTime.Now.AddDays(-5);
            new Program(new DateTime(temp.Year, temp.Month, temp.Day, 0, 0, 0));
        }
    }
}
