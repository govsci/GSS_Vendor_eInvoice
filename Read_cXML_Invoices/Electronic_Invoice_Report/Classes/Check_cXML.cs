using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electronic_Invoice_Report.Objects;
using System.IO;
using System.Xml;

namespace Electronic_Invoice_Report.Classes
{
    public static class Check_cXML
    {
        private delegate string F1(XmlNode x);

        private static List<Invoice> invoices = new List<Invoice>();
        public static List<Invoice> Check(DateTime check)
        {
            invoices = new List<Invoice>();
            for (DateTime day = check.Date; day <= DateTime.Now.Date; day = day.AddDays(1))
                GetFiles(day);
            return invoices;
        }

        private static void GetFiles(DateTime day)
        {
            string folder = Constants.cXMLInvoiceFolder + $@"{day.ToString("yyyy")}\{day.ToString("MM")}\{day.ToString("dd")}\";
            string[] files = Directory.GetFiles(folder);
            foreach(string file in files)
            {
                XmlDocument xml = new XmlDocument();
                xml.Load(file);
                ReadFile(file, xml);
            }
        }

        private static void ReadFile(string file, XmlDocument xml)
        {
            F1 setValue = x => x == null ? "" : x.InnerXml;

            string vendor = Database.GetVendor(
                setValue(xml.SelectSingleNode("//Header/Sender/UserAgent")),
                setValue(xml.SelectSingleNode("//Header/From/Credential/Identity")),
                setValue(xml.SelectSingleNode("//Header/Sender/Credential/SharedSecret")),
                setValue(xml.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceID")),
                setValue(xml.SelectSingleNode("//InvoiceDetailRequestHeader/InvoicePartner/Contact[@role='remitTo']/Name")));

            invoices.Add(new Invoice(
                vendor,
                "XML",
                setValue(xml.SelectSingleNode("//Header/From/Credential/@domain")),
                setValue(xml.SelectSingleNode("//Header/From/Credential/Identity")),
                setValue(xml.SelectSingleNode("//Header/To/Credential/@domain")),
                setValue(xml.SelectSingleNode("//Header/To/Credential/Identity")),
                setValue(xml.SelectSingleNode("//Header/Sender/Credential/@domain")),
                setValue(xml.SelectSingleNode("//Header/Sender/Credential/Identity")),
                setValue(xml.SelectSingleNode("//Header/Sender/Credential/SharedSecret")),
                setValue(xml.SelectSingleNode("//Header/Sender/UserAgent")),
                setValue(xml.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceID")), 
                setValue(xml.SelectSingleNode("//OrderReference/@orderID")),
                File.GetCreationTime(file),
                file
                ));
        }
    }
}
