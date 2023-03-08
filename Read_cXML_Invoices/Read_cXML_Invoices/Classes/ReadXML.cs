using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes
{
    public class ReadXML
    {
        private string file = "";
        private InvoiceHeader invoice;

        public delegate string F1(XmlNode x);
        public delegate decimal F2(XmlNode x);
        public delegate int F3(XmlNode x);

        public ReadXML(string f)
        {
            file = f;
            invoice = new InvoiceHeader();
            ExtractXmlNEW();
        }
        public InvoiceHeader Invoice { get { return invoice; } }

        private bool CheckPrices(decimal quantity, decimal unitPrice, decimal lineTotal)
        {
            try
            {
                decimal test = unitPrice * quantity;
                if (lineTotal == test)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }

        private bool CheckInvoice(string invoiceid, string orderid)
        {
            bool rVal = false;
            using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
            {
                dbcon.Open();
                string query = $"SELECT [invoiceID] FROM [dbo].[Ecommerce$Electronic Document Header] WHERE [documentType] = 'Invoice' AND [invoiceID]='{invoiceid}' AND [orderID] = '{orderid}'";
                using (SqlDataReader rs = new SqlCommand(query, dbcon).ExecuteReader())
                    if (rs.Read())
                        rVal = true;
            }
            return rVal;
        }

        private void ExtractXmlNEW()
        {
            try
            {
                decimal dtest = 0.0M;
                int itest = 0;
                F1 setValue = x => x == null ? "" : x.InnerXml.Replace("'", "''");
                F2 setDecimalValue = x => x != null && decimal.TryParse(x.InnerXml, out dtest) ? decimal.Parse(x.InnerXml) : 0.0M;
                F3 setIntValue = x => x != null && int.TryParse(x.InnerXml, out itest) ? int.Parse(x.InnerXml) : 0;

                XmlDocument xml = new XmlDocument();
                xml.XmlResolver = null;
                string xmltext = File.ReadAllText(file);
                if (xmltext.Contains("cXML_Invoice")) xmltext = xmltext.Replace("cXML_Invoice", "cXML");
                if (xmltext.Contains("xmlns:ns0=\"http://www.w3.org/XML/1998/namespace\"")) xmltext = xmltext.Replace("xmlns:ns0=\"http://www.w3.org/XML/1998/namespace\"", "");
                xml.LoadXml(xmltext);

                string invoiceid = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/@invoiceID"));
                string orderid = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailOrder/InvoiceDetailOrderInfo/OrderReference/@orderID"));
                if (orderid.Length == 0) orderid = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailOrder/InvoiceDetailOrderInfo/OrderIDInfo/@orderID"));
                if (CheckInvoice(invoiceid, orderid))
                {
                    invoice = null;
                    return;
                }

                invoice = new InvoiceHeader();

                //cXML Document Data
                invoice.Timestamp = setValue(xml.SelectSingleNode("cXML/@timestamp"));
                invoice.PayloadID = setValue(xml.SelectSingleNode("cXML/@payloadID"));

                //cXML Header
                invoice.FromDomain = setValue(xml.SelectSingleNode("cXML/Header/From/Credential/@domain"));
                invoice.FromIdentity = setValue(xml.SelectSingleNode("cXML/Header/From/Credential/Identity"));
                invoice.ToDomain = setValue(xml.SelectSingleNode("cXML/Header/To/Credential/@domain"));
                invoice.ToIdentity = setValue(xml.SelectSingleNode("cXML/Header/To/Credential/Identity"));
                invoice.SenderDomain = setValue(xml.SelectSingleNode("cXML/Header/Sender/Credential/@domain"));
                invoice.SenderIdentity = setValue(xml.SelectSingleNode("cXML/Header/Sender/Credential/Identity"));
                invoice.SharedSecret = setValue(xml.SelectSingleNode("cXML/Header/Sender/Credential/SharedSecret"));
                invoice.UserAgent = setValue(xml.SelectSingleNode("cXML/Header/Sender/UserAgent"));

                //cXML InvoiceDetailRequestHeader
                invoice.InvoiceID = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/@invoiceID"));
                invoice.Purpose = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/@purpose"));
                invoice.Operation = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/@operation"));
                invoice.InvoiceDate = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/@invoiceDate"));

                if (invoice.UserAgent == "" && (invoice.InvoiceID.IndexOf("IN-") > -1 || invoice.InvoiceID.IndexOf("CM-") > -1)) invoice.UserAgent = "OfficeCity";
                if (invoice.UserAgent == "" && invoice.InvoiceID.IndexOf("VANE") > -1) invoice.UserAgent = "Fastenal";

                XmlNodeList invoicePartners = xml.SelectNodes("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/InvoicePartner");
                foreach (XmlNode invoicePartner in invoicePartners)
                {
                    string role = setValue(invoicePartner.SelectSingleNode("Contact/@role"));
                    string addressID = setValue(invoicePartner.SelectSingleNode("Contact/@addressID"));
                    string name = setValue(invoicePartner.SelectSingleNode("Contact/Name"));
                    if (name.Length == 0) name = setValue(invoicePartner.SelectSingleNode("Contact/PostalAddress/@name"));
                    string deliverTo = setValue(invoicePartner.SelectSingleNode("Contact/PostalAddress/DeliverTo"));

                    string street = "";
                    foreach (XmlNode node in invoicePartner.SelectNodes("Contact/PostalAddress/Street"))
                        if (setValue(node).Length > 0)
                            street = street.Length > 0 ? street + "|" + setValue(node) : setValue(node);

                    string city = setValue(invoicePartner.SelectSingleNode("Contact/PostalAddress/City"));
                    string state = setValue(invoicePartner.SelectSingleNode("Contact/PostalAddress/State"));
                    string postalCode = setValue(invoicePartner.SelectSingleNode("Contact/PostalAddress/PostalCode"));
                    string countryCode = setValue(invoicePartner.SelectSingleNode("Contact/PostalAddress/Country/@isoCountryCode"));
                    string country = setValue(invoicePartner.SelectSingleNode("Contact/PostalAddress/Country"));

                    invoice.Roles.Add(new AddressObject(role
                        , addressID
                        , name
                        , deliverTo
                        , street
                        , city
                        , state
                        , postalCode
                        , countryCode
                        , country));
                }

                XmlNodeList invoiceDetailShippings = xml.SelectNodes("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/InvoiceDetailShipping/Contact");
                foreach (XmlNode invoiceDetailShipping in invoiceDetailShippings)
                {
                    string role = setValue(invoiceDetailShipping.SelectSingleNode("@role"));
                    string addressID = setValue(invoiceDetailShipping.SelectSingleNode("@addressID"));
                    string name = setValue(invoiceDetailShipping.SelectSingleNode("Name"));
                    if (name.Length == 0) name = setValue(invoiceDetailShipping.SelectSingleNode("PostalAddress/@name"));
                    string deliverTo = setValue(invoiceDetailShipping.SelectSingleNode("PostalAddress/DeliverTo"));

                    string street = "";
                    foreach (XmlNode node in invoiceDetailShipping.SelectNodes("PostalAddress/Street"))
                        if (setValue(node).Length > 0)
                            street = street.Length > 0 ? street + "|" + setValue(node) : setValue(node);

                    string city = setValue(invoiceDetailShipping.SelectSingleNode("PostalAddress/City"));
                    string state = setValue(invoiceDetailShipping.SelectSingleNode("PostalAddress/State"));
                    string postalCode = setValue(invoiceDetailShipping.SelectSingleNode("PostalAddress/PostalCode"));
                    string countryCode = setValue(invoiceDetailShipping.SelectSingleNode("PostalAddress/Country/@isoCountryCode"));
                    string country = setValue(invoiceDetailShipping.SelectSingleNode("PostalAddress/Country"));

                    invoice.Roles.Add(new AddressObject(role
                        , addressID
                        , name
                        , deliverTo
                        , street
                        , city
                        , state
                        , postalCode
                        , countryCode
                        , country));
                }

                invoice.PaymentTermNumberOfDays = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/InvoiceDetailPaymentTerm/@payInNumberOfDays"));
                invoice.PaymentTermPercentRate = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/InvoiceDetailPaymentTerm/@percentageRate"));

                XmlNodeList extrinsics = xml.SelectNodes("cXML/Request/InvoiceDetailRequest/InvoiceDetailRequestHeader/Extrinsic");
                foreach (XmlNode extrinsic in extrinsics)
                    invoice.Extrinsics.Add(new Extrinsic(setValue(extrinsic.SelectSingleNode("@name")), setValue(extrinsic)));

                //cXML InvoiceDetailOrder
                invoice.OrderID = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailOrder/InvoiceDetailOrderInfo/OrderReference/@orderID"));
                if (invoice.OrderID.Length == 0) invoice.OrderID = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailOrder/InvoiceDetailOrderInfo/OrderIDInfo/@orderID"));

                invoice.OrderDate = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailOrder/InvoiceDetailOrderInfo/OrderReference/@orderDate"));
                invoice.DocumentReferencePayloadID = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailOrder/InvoiceDetailOrderInfo/OrderReference/DocumentReference/@payloadID"));

                if (invoice.Roles.Find(r => r.Role == "shipTo") == null)
                {
                    AddressObject address = Database.GetShippingInformation(invoice.OrderID);
                    if (address != null)
                        invoice.Roles.Add(address);
                }

                invoice.InvoiceTotal = 0.0M;
                decimal shippingTotal = 0.0M;
                XmlNodeList invoiceDetailItems = xml.SelectNodes("cXML/Request/InvoiceDetailRequest/InvoiceDetailOrder/InvoiceDetailItem");
                foreach (XmlNode invoiceDetailItem in invoiceDetailItems)
                {
                    InvoiceLine line = new InvoiceLine();
                    line.LineNumber = setIntValue(invoiceDetailItem.SelectSingleNode("@invoiceLineNumber"));
                    line.Quantity = setDecimalValue(invoiceDetailItem.SelectSingleNode("@quantity"));
                    line.UnitOfMeasure = setValue(invoiceDetailItem.SelectSingleNode("UnitOfMeasure"));
                    line.UnitPrice = setDecimalValue(invoiceDetailItem.SelectSingleNode("UnitPrice/Money"));
                    line.ReferenceLineNumber = setIntValue(invoiceDetailItem.SelectSingleNode("InvoiceDetailItemReference/@lineNumber"));
                    line.SupplierPartID = setValue(invoiceDetailItem.SelectSingleNode("InvoiceDetailItemReference/ItemID/SupplierPartID"));
                    line.Description = setValue(invoiceDetailItem.SelectSingleNode("InvoiceDetailItemReference/Description"));
                    line.Tax = setDecimalValue(invoiceDetailItem.SelectSingleNode("Tax/Money"));
                    line.LineTotal = line.Quantity * line.UnitPrice;
                    if (line.UnitPrice == 0.0M) line.LineTotal = line.UnitPrice = setDecimalValue(invoiceDetailItem.SelectSingleNode("SubtotalAmount/Money"));
                    invoice.Lines.Add(line);

                    if (!line.SupplierPartID.ToUpper().Contains("SHIP"))
                        invoice.InvoiceTotal += line.LineTotal;
                    else
                        shippingTotal += line.LineTotal;

                    if (invoiceDetailItem.SelectSingleNode("InvoiceDetailLineShipping") != null)
                    {
                        XmlNodeList shipping = invoiceDetailItem.SelectNodes("InvoiceDetailLineShipping/InvoiceDetailShipping/Contact");
                        foreach (XmlNode ship in shipping)
                        {
                            string role = setValue(ship.SelectSingleNode("@role"));
                            string addressID = setValue(ship.SelectSingleNode("@addressID"));
                            if (invoice.Roles.Find(r => r.Role == role && r.ID == addressID) == null)
                            {
                                invoice.Roles.Add(new AddressObject(role
                                    , addressID
                                    , setValue(ship.SelectSingleNode("Name"))
                                    , setValue(ship.SelectSingleNode("PostalAddress/DeliverTo"))
                                    , setValue(ship.SelectSingleNode("PostalAddress/Street"))
                                    , setValue(ship.SelectSingleNode("PostalAddress/City"))
                                    , setValue(ship.SelectSingleNode("PostalAddress/State"))
                                    , setValue(ship.SelectSingleNode("PostalAddress/PostalCode"))
                                    , setValue(ship.SelectSingleNode("PostalAddress/Country/@isoCountryCode"))
                                    , setValue(ship.SelectSingleNode("PostalAddress/Country"))));
                            }
                        }
                    }
                }

                invoice.SubTotalAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/SubtotalAmount/Money"));
                if (invoice.SubTotalAmount == 0.0M) invoice.SubTotalAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/SubtotalAmount"));
                if (invoice.InvoiceTotal > invoice.SubTotalAmount) invoice.SubTotalAmount = invoice.InvoiceTotal;

                string taxDescription = setValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/Tax/Description"));
                if (taxDescription.StartsWith("EDI"))
                {
                    invoice.SpecialHandlingAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/Tax/Money"));
                    if (invoice.SpecialHandlingAmount == 0.0M) setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/Tax"));
                }
                else
                {
                    invoice.Tax = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/Tax/Money"));
                    if (invoice.Tax == 0.0M) setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/Tax"));
                }

                if (xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/ShippingAmount") != null)
                {
                    decimal shippingtotal = 0.0M;
                    XmlNodeList shippingNodes = xml.SelectNodes("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/ShippingAmount");
                    for (int i = 0; i < shippingNodes.Count; i++)
                    {
                        int index = i + 1;
                        decimal shipping = setDecimalValue(shippingNodes[i].SelectSingleNode("Money"));
                        if (shipping > 0.0m)
                        {
                            Extrinsic extr = invoice.Extrinsics.Find(e => e.Name == $"shippingAmount{index}");
                            if (extr != null)
                            {
                                InvoiceLine line = new InvoiceLine();
                                line.LineNumber = 0;
                                line.Quantity = 1;
                                line.UnitOfMeasure = "EA";
                                line.UnitPrice = shipping;
                                line.ReferenceLineNumber = 0;
                                line.SupplierPartID = extr.Value;
                                line.Description = extr.Value;
                                line.Tax = 0;
                                line.LineTotal = line.Quantity * line.UnitPrice;
                                line.ShipLine = 1;
                                invoice.Lines.Add(line);
                            }
                            else
                                shippingtotal += shipping;
                        }
                        if (shipping == 0.0M)
                            shippingtotal += setDecimalValue(shippingNodes[i]);
                    }

                    invoice.ShippingAmount = shippingtotal;
                }

                invoice.NetAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/NetAmount/Money"));
                if (invoice.NetAmount == 0.0M) invoice.NetAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/NetAmount"));

                invoice.GrossAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/GrossAmount/Money"));
                if (invoice.GrossAmount == 0.0M) invoice.GrossAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/GrossAmount"));

                invoice.SpecialHandlingAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/SpecialHandlingAmount/Money"));
                if (invoice.SpecialHandlingAmount == 0.0M) invoice.SpecialHandlingAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/SpecialHandlingAmount"));

                invoice.DueAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/DueAmount/Money"));
                if (invoice.DueAmount == 0.0M) invoice.DueAmount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/DueAmount"));

                if (shippingTotal == invoice.SpecialHandlingAmount) invoice.InvoiceTotal += invoice.Tax + invoice.ShippingAmount + shippingTotal;
                else invoice.InvoiceTotal += invoice.Tax + invoice.ShippingAmount + invoice.SpecialHandlingAmount + shippingTotal;

                invoice.InvoiceDetailDiscount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/InvoiceDetailDiscount/Money"));
                if (invoice.InvoiceDetailDiscount == 0.0M) invoice.InvoiceDetailDiscount = setDecimalValue(xml.SelectSingleNode("cXML/Request/InvoiceDetailRequest/InvoiceDetailSummary/InvoiceDetailDiscount"));

                AddressObject remitTo = invoice.Roles.Find(a => a.Role == "remitTo");
                Database.GetVendor(ref invoice, remitTo);

                invoice.ReceiveDate = DateTime.Now;
                string[] poarray = invoice.OrderID.Split('/');
                string po = "";
                if (poarray.Length > 1)
                    po = poarray[1];
                else
                    po = poarray[0];
            }
            catch(Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, "ReadXML", "ExtractXmlNEW"));
            }
        }        
    }
}
