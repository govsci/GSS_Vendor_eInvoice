using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Read_cXML_Invoices.Objects;
using System.IO;
using System.Xml;

namespace Read_cXML_Invoices.Classes
{
    public class Check_cXML
    {
        private delegate string F1(XmlNode x);

        private List<Invoice> invoices = new List<Invoice>();
        private DateTime AppStarted;
        public List<Invoice> Check(DateTime check, DateTime appStarted)
        {
            AppStarted = appStarted;
            invoices = new List<Invoice>();
            for (DateTime day = check.Date; day <= DateTime.Now.Date; day = day.AddDays(1))
                GetFiles(day);
            return invoices;
        }

        private void GetFiles(DateTime day)
        {
            string folder = Constants.InvoiceFolder + $@"{day.ToString("yyyy")}\{day.ToString("MM")}\{day.ToString("dd")}\";
            if (Directory.Exists(folder))
            {
                string[] files = Directory.GetFiles(folder);
                foreach (string file in files)
                {
                    if (File.GetCreationTime(file) < AppStarted)
                    {
                        XmlDocument xml = new XmlDocument();
                        xml.Load(file);
                        ReadFile(file, xml);
                    }
                }
            }
        }

        private void ReadFile(string file, XmlDocument xml)
        {
            F1 setValue = x => x == null ? "" : x.InnerXml;

            string remitToName = setValue(xml.SelectSingleNode("//InvoiceDetailRequestHeader/InvoicePartner/Contact[@role='remitTo']/Name"));
            if (remitToName.Length == 0)
                remitToName = setValue(xml.SelectSingleNode("//InvoiceDetailRequestHeader/InvoicePartner/Contact[@role='remitTo']/PostalAddress/@name"));

            string vendor = Database.GetVendor(
                setValue(xml.SelectSingleNode("//Header/Sender/UserAgent")),
                setValue(xml.SelectSingleNode("//Header/From/Credential/Identity")),
                setValue(xml.SelectSingleNode("//Header/Sender/Credential/SharedSecret")),
                setValue(xml.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceID")),
                remitToName);

            string checkTable = Database.CheckInvoice(setValue(xml.SelectSingleNode("//InvoiceDetailRequestHeader/@invoiceID")), setValue(xml.SelectSingleNode("//OrderReference/@orderID")));

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
                file,
                checkTable == "NOT_FOUND" ? "NO" : "Yes",
                checkTable != "0" && checkTable != "NOT_FOUND" ? checkTable : ""
                ));
        }
    }
}
