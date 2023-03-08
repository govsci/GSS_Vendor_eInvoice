using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing.Layout;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.CompilerServices;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes
{
    public class BuildSingleInvoicePDF
    {
        private string saveFile = "";
        private string masterSaveFile = "";
        private string backupSaveFile = "";
        private string backupMasterSaveFile = "";
        private InvoiceHeader invoice;
        private List<string> errorList = new List<string>();
        public string TimeStamp(DateTime value) { return value.ToString("yyyyMMddHHmmssffff"); }

        public BuildSingleInvoicePDF(InvoiceHeader i, string subFolder, string mainFolder)
        {
            invoice = i;
            saveFile = subFolder + @"\";
            backupSaveFile = "" + saveFile;
            masterSaveFile = mainFolder + @"\";
            backupMasterSaveFile = "" + masterSaveFile;
        }

        public string CreatePDF()
        {
            PdfDocument docu = new PdfDocument();
            docu.Info.Title = DateTime.Now.ToShortDateString() + " INVOICES";
            //XFont fontBold = new XFont("Arial", 9, XFontStyle.Bold);
            XGraphics gfx = null;
            XFont font = null;
            XTextFormatter tf = null;
            XRect rect;

            int pageLineCount = 6, lineStartPositionY = 260;
            int offsetX = 30, boxPadding = 5, itemLineHeight = 0, pageLineCounter = 0, pageCount = 0;

            decimal previousPageTotal = 0, currentPageTotal = 0;
            if (invoice.Lines.Count == 0)
            {
                XPen pen = new XPen(XColors.Gray, 0.5);
                pen.DashStyle = XDashStyle.Solid;
                XGraphicsPath path = new XGraphicsPath();
                if (pageLineCounter % pageLineCount == 0)
                {
                    if (pageCount > 0)
                    {
                        //rect = new XRect(offsetX + 80 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 600 - boxPadding, 75 - boxPadding);
                        //tf.DrawString("Continued on page " + ((pageCount + 1).ToString()) + ".....................................................................................................................", font, XBrushes.Black, rect, XStringFormats.TopLeft);
                        //tf.Alignment = XParagraphAlignment.Right;
                        //rect = new XRect(offsetX + 290 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                        //tf.DrawString(currentPageTotal.ToString("G29"), font, XBrushes.Black, rect, XStringFormats.TopLeft);
                        itemLineHeight += 40;
                        tf.Alignment = XParagraphAlignment.Left;
                        previousPageTotal = currentPageTotal;
                    }
                    gfx = addPdfPage(docu, invoice);
                    pageCount++;
                    pageLineCounter = 0;
                    itemLineHeight = 0;
                }
                font = new XFont("Arial", 10);
                tf = new XTextFormatter(gfx);
            }
            else
            {
                for (int j = 0; j < invoice.Lines.Count; j++)
                {
                    XPen pen = new XPen(XColors.Gray, 0.5);
                    pen.DashStyle = XDashStyle.Solid;
                    XGraphicsPath path = new XGraphicsPath();
                    if (pageLineCounter % pageLineCount == 0)
                    {
                        if (pageCount > 0)
                        {
                            //rect = new XRect(offsetX + 80 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 600 - boxPadding, 75 - boxPadding);
                            //tf.DrawString("Continued on page " + ((pageCount + 1).ToString()) + ".....................................................................................................................", font, XBrushes.Black, rect, XStringFormats.TopLeft);
                            //tf.Alignment = XParagraphAlignment.Right;
                            //rect = new XRect(offsetX + 290 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                            //tf.DrawString(currentPageTotal.ToString("G29"), font, XBrushes.Black, rect, XStringFormats.TopLeft);
                            itemLineHeight += 40;
                            tf.Alignment = XParagraphAlignment.Left;
                            previousPageTotal = currentPageTotal;
                        }
                        gfx = addPdfPage(docu, invoice);
                        pageCount++;
                        pageLineCounter = 0;
                        itemLineHeight = 0;
                    }
                    currentPageTotal += invoice.Lines[j].UnitPrice;
                    font = new XFont("Arial", 10);
                    tf = new XTextFormatter(gfx);
                    if (pageCount > 1 && pageLineCounter == 0)
                    {
                        //rect = new XRect(offsetX + 80 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 600 - boxPadding, 75 - boxPadding);
                        //tf.DrawString("Continued from page " + ((pageCount - 1).ToString()) + "..................................................................................................................", font, XBrushes.Black, rect, XStringFormats.TopLeft);
                        //string totalString = String.Format("{0:0,0.00}", previousPageTotal);
                        //tf.Alignment = XParagraphAlignment.Right;
                        //rect = new XRect(offsetX + 290 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                        //tf.DrawString(totalString, font, XBrushes.Black, rect, XStringFormats.TopLeft);
                        itemLineHeight += 40;
                        tf.Alignment = XParagraphAlignment.Left;
                    }

                    rect = new XRect(offsetX + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                    tf.DrawString(invoice.Lines[j].LineNumber.ToString(), font, XBrushes.Black, rect, XStringFormats.TopLeft);
                    //tf.DrawString(invoice.Lines[j].SupplierPartID, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    rect = new XRect(offsetX + 35 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                    tf.DrawString(invoice.Lines[j].SupplierPartID, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    rect = new XRect(offsetX + 120 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                    tf.DrawString(invoice.Lines[j].UnitOfMeasure, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    string desc = invoice.Lines[j].Description;
                    if (desc.Length > 100)
                        desc = desc.Substring(0, 100) + "...";
                    rect = new XRect(offsetX + 150 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                    tf.DrawString(desc, font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    rect = new XRect(offsetX + 386 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                    tf.DrawString(invoice.Lines[j].Quantity.ToString("0"), font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    rect = new XRect(offsetX + 435 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                    tf.DrawString(invoice.Lines[j].UnitPrice.ToString("G29"), font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    rect = new XRect(offsetX + 490 + boxPadding, lineStartPositionY + boxPadding + itemLineHeight, 230 - boxPadding, 75 - boxPadding);
                    tf.DrawString(invoice.Lines[j].LineTotal.ToString("G29"), font, XBrushes.Black, rect, XStringFormats.TopLeft);

                    /*if (desc.Length > 205)
                        itemLineHeight += 65; else */
                    if (desc.Length > 185)
                        itemLineHeight += 60;
                    else if (desc.Length > 96)
                        itemLineHeight += 50;
                    else
                        itemLineHeight += 40;
                    pageLineCounter++;
                }
            }
            if (gfx == null)
            {
                errorList.Add("\r\nAdding a PDF failed:\r\n" + invoice.InvoiceID + "\r\n");
                return "";
            }
            XPen pen2 = new XPen(XColors.Black, 0.5);
            pen2.DashStyle = XDashStyle.Solid;
            XGraphicsPath path2 = new XGraphicsPath();
            path2.AddLine(offsetX, 680, 570, 680);
            gfx.DrawPath(pen2, path2);
            XFont fontNormal = new XFont("Arial", 10); //, XFontStyle.Bold);
            //XFont boldFont = new XFont("Arial", 9, XFontStyle.Bold);
            XFont headerFont = new XFont("Times New Roman", 20, XFontStyle.BoldItalic);

            XFont italicFont = new XFont("Helvetica", 9, XFontStyle.Italic);
            string mid = "This is a receipt of all items contained in this package. This invoice represents a communication\n"
                + "received by " + invoice.Vendor + " double punchout system, this invoice doesn't contain any shipping \n"
                + "charge from " + invoice.Vendor + ".";
            rect = new XRect(offsetX + 5, 690, 540, 232);
            tf.DrawString(mid, italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 220, 725, 540, 232);
            tf.DrawString("Thank You", headerFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            string lineShippingCaption = "";
            string lineShipping = "";
            decimal shipLineTotal = 0.0m;
            foreach (InvoiceLine shipline in invoice.ShipLines)
            {
                shipLineTotal += shipline.LineTotal;
                lineShippingCaption += ReturnShortenedShippingDescription(shipline.Description.Trim().Length > 0 ? shipline.Description.Trim() : shipline.SupplierPartID) + ":\n";
                lineShipping += shipline.LineTotal.ToString("G29") + "\n";
            }
            if (shipLineTotal == invoice.SpecialHandlingAmount) invoice.SpecialHandlingAmount = 0.0m;

            if (invoice.SpecialHandlingAmount > 0.0M)
            {
                lineShippingCaption += "Special Handling:\n";
                lineShipping += invoice.SpecialHandlingAmount.ToString("G29") + "\n";
            }

            if (invoice.InvoiceDetailDiscount != 0.0M)
            {
                lineShippingCaption += "Discount:\n";
                lineShipping += invoice.InvoiceDetailDiscount > 0.0M ? "-" + invoice.InvoiceDetailDiscount.ToString("G29") + "\n" : invoice.InvoiceDetailDiscount.ToString("G29") + "\n";

                if (invoice.InvoiceDetailDiscount != 0.0M && invoice.SubTotalAmount == invoice.InvoiceTotal) 
                {
                    if (invoice.InvoiceDetailDiscount > 0.0M)
                        invoice.InvoiceTotal = invoice.InvoiceTotal - invoice.InvoiceDetailDiscount;
                    else if (invoice.InvoiceDetailDiscount < 0.0M)
                        invoice.InvoiceTotal = invoice.InvoiceTotal + invoice.InvoiceDetailDiscount; 
                }
            }

            rect = new XRect(offsetX + 400, 690, 540, 232);
            tf.DrawString("Subtotal:\nShipping:\n" + lineShippingCaption + "Tax:\nTotal:\n", fontNormal, XBrushes.Black, rect, XStringFormats.TopLeft);

            string totals = invoice.SubTotalAmount.ToString("G29") + "\n";
            totals += invoice.ShippingAmount.ToString("G29") + "\n"
                + lineShipping
                + invoice.Tax.ToString("G29") + "\n"
                + invoice.InvoiceTotal.ToString("G29") + "\n";

            rect = new XRect(offsetX + 500, 690, 540, 232);
            tf.DrawString(totals, fontNormal, XBrushes.Black, rect, XStringFormats.TopLeft);

            try
            {
                docu.Save(saveFile);
                docu.Save(masterSaveFile);
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, new System.Data.SqlClient.SqlCommand(saveFile), "BuildSingleInvoicePDF", "CreatePDF"));
                string fileName = DateTime.Now.ToString("yyyyMMddThhmmssfffffff") + ".pdf";
                saveFile = backupSaveFile + fileName;
                masterSaveFile = backupMasterSaveFile + fileName;
                docu.Save(saveFile);
                docu.Save(masterSaveFile);
            }
            return saveFile;
        }
        private XGraphics addPdfPage(PdfDocument document, InvoiceHeader data)
        {
            string header = "", header2 = "";
            AddressObject remitTo = data.Roles.Find(r => r.Role == "remitTo");
            AddressObject issuerOfInvoice = data.Roles.Find(r => r.Role == "issuerOfInvoice" || r.Role == "IssuerOfInvoice");
            AddressObject shipTo = data.Roles.Find(r => r.Role == "shipTo");
            AddressObject billTo = data.Roles.Find(r => r.Role == "billTo");
            AddressObject soldTo = data.Roles.Find(r => r.Role == "soldTo");

            if (remitTo != null && remitTo.Name.Length > 0)
            {
                header = "Invoiced: " + remitTo.Name;
                if (remitTo.Street.Length > 0) header2 += remitTo.Street.Replace("|", "\n") + "\n";
                if (remitTo.City.Length > 0) header2 += remitTo.City + ", " + remitTo.State + " " + remitTo.PostalCode;
                if (remitTo.Country.Length > 0) header2 += "\n" + remitTo.Country;
            }
            else if (issuerOfInvoice != null && issuerOfInvoice.Name.Length > 0)
            {
                header = "Invoiced: " + issuerOfInvoice.Name;
                if (issuerOfInvoice.Street.Length > 0) header2 += issuerOfInvoice.Street.Replace("|", "\n") + "\n";
                if (issuerOfInvoice.City.Length > 0) header2 += issuerOfInvoice.City + ", " + issuerOfInvoice.State + " " + issuerOfInvoice.PostalCode;
                if (issuerOfInvoice.Country.Length > 0) header2 += "\n" + issuerOfInvoice.Country;
            }
            else
            {
                header = "Invoiced:";
            }

            if (!saveFile.Contains(".pdf"))
            {
                string fileName = data.InvoiceID + "." + data.OrderID + "." + data.Vendor + " " + invoice.ShipType + ".pdf";

                saveFile += fileName.Replace(@"\", "").Replace("*", "").Replace(@"\", "").Replace("/", "").Replace(":", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "");
                masterSaveFile += fileName.Replace(@"\", "").Replace("*", "").Replace(@"\", "").Replace("/", "").Replace(":", "").Replace("?", "").Replace("\"", "").Replace("<", "").Replace(">", "").Replace("|", "");
            }

            string order = data.Vendor + " INVOICE";
            if (data.DueAmount < 0)
                order += "\nCREDIT MEMO";

            string ship = "";
            if (shipTo != null)
            {
                ship = shipTo.Name;
                if (shipTo.Name.Length > 0) ship += "\n";
                ship += shipTo.Street.Replace("|", "\n");
                if (shipTo.Street.Length > 0) ship += "\n";
                if (shipTo.City.Length > 0)
                    ship += shipTo.City + ", " + shipTo.State + " " + shipTo.PostalCode;
                if (shipTo.Country.Length > 0) ship += "\n" + shipTo.Country;
            }

            string bill = "";
            if (billTo != null && (billTo.Name.Length > 0 || billTo.Street.Length > 0))
            {
                bill = billTo.Name;
                if (!billTo.Name.Equals("")) bill += "\n";
                bill += billTo.Street.Replace("|", "\n");
                if (!billTo.Street.Equals("")) bill += "\n";
                if (billTo.City.Length > 0)
                    bill += billTo.City + ", " + billTo.State + " " + billTo.PostalCode;
                if (billTo.Country.Length > 0) bill += "\n" + billTo.Country;
            }
            else if (soldTo != null && (soldTo.Name.Length > 0 || soldTo.Street.Length > 0))
            {
                bill = soldTo.Name;
                if (!soldTo.Name.Equals("")) bill += "\n";
                bill += soldTo.Street.Replace("|", "\n");
                if (!soldTo.Street.Equals("")) bill += "\n";
                if (soldTo.City.Length > 0)
                    bill += soldTo.City + ", " + soldTo.State + " " + soldTo.PostalCode;
                if (soldTo.Country.Length > 0) bill += "\n" + soldTo.Country;
            }
            else
                bill = "Government Scientific Source\n12351 Sunrise Valley Drive\nReston, VA 20191\nUS";

            int offsetX = 30;
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            XFont font = new XFont("Arial", 10); //, XFontStyle.Bold);
            //XFont boldFont = new XFont("Arial", 9, XFontStyle.Bold);
            XFont headerFont = new XFont("Times New Roman", 13, XFontStyle.BoldItalic);
            XFont invoiceFont = new XFont("Arial", 13, XFontStyle.Bold);
            XFont italicFont = new XFont("Helvetica", 9, XFontStyle.Italic);
            XTextFormatter tf = new XTextFormatter(gfx);

            //Header
            XRect rect = new XRect(offsetX, 30, 540, 232);
            tf.DrawString(header, headerFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 8, 51, 540, 232);
            tf.DrawString(header2, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 350, 30, 540, 232);
            tf.DrawString(order, invoiceFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 400, 65, 540, 232);
            tf.DrawString("P.O. #:", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 460, 65, 540, 232);
            tf.DrawString(data.OrderID, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 400, 80, 540, 232);
            tf.DrawString("Invoice #:", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 460, 80, 540, 232);
            tf.DrawString(data.InvoiceID, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 400, 95, 540, 232);
            tf.DrawString("Invoice Date:", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 460, 95, 540, 232);
            DateTime test = new DateTime();
            if (DateTime.TryParse(data.InvoiceDate.Replace("T", " "), out test))
                tf.DrawString(Convert.ToDateTime(data.InvoiceDate.Replace("T", " ")).ToShortDateString(), font, XBrushes.Black, rect, XStringFormats.TopLeft);
            else
                tf.DrawString("(Invoice Date Not Available)", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 400, 110, 540, 232);
            tf.DrawString("Pay Terms:", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 460, 110, 540, 232);
            tf.DrawString(data.PaymentTermNumberOfDays + " days", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 8, 110, 540, 232);
            tf.DrawString("Ship To:", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 8, 125, 540, 232);
            tf.DrawString(ship, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 200, 110, 540, 232);
            tf.DrawString("Bill To:", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 200, 125, 540, 232);
            tf.DrawString(bill, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            string extrinsics = "";
            foreach (Extrinsic ext in data.Extrinsics)
                extrinsics += ext.Name + ": " + ext.Value + "\n";

            rect = new XRect(offsetX + 392, 125, 540, 232);
            tf.DrawString(extrinsics, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            //Middle - First
            XPen pen = new XPen(XColors.Gray, 0.5);
            pen.DashStyle = XDashStyle.Solid;
            XGraphicsPath path = new XGraphicsPath();
            path.AddLine(offsetX, 200, 570, 200);
            gfx.DrawPath(pen, path);

            rect = new XRect(offsetX + 15, 202, 540, 232);
            tf.DrawString("Order Date", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 150, 202, 540, 232);
            tf.DrawString("Invoice Received", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 270, 202, 540, 232);
            tf.DrawString("Tracking #", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 425, 202, 540, 232);
            tf.DrawString("Account #", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            XPen pen2 = new XPen(XColors.Gray, 0.5);
            pen2.DashStyle = XDashStyle.Solid;
            XGraphicsPath path2 = new XGraphicsPath();
            path2.AddLine(offsetX, 212, 570, 212);
            gfx.DrawPath(pen2, path2);

            string orderdatestring = data.OrderDate;
            test = DateTime.Now;
            if (DateTime.TryParse(orderdatestring, out test))
                orderdatestring = DateTime.Parse(data.OrderDate).ToShortDateString();
            rect = new XRect(offsetX + 15, 215, 540, 232);
            tf.DrawString(orderdatestring, font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 150, 215, 540, 232);
            tf.DrawString(data.ReceiveDate.ToShortDateString(), font, XBrushes.Black, rect, XStringFormats.TopLeft);

            Extrinsic trackingNumber = data.Extrinsics.Find(e => e.Name.ToUpper() == "TRACKINGNO");
            if (trackingNumber == null) trackingNumber = data.Extrinsics.Find(e => e.Name.ToUpper() == "CFVALUE_ORDER_REFERENCE_NUMBER");
            Extrinsic acctNumber = data.Extrinsics.Find(e => e.Name.ToUpper() == "ACCOUNTNO");

            rect = new XRect(offsetX + 270, 215, 540, 232);
            tf.DrawString(trackingNumber != null ? trackingNumber.Value : "", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 430, 215, 540, 232);
            tf.DrawString(acctNumber != null ? acctNumber.Value : "", font, XBrushes.Black, rect, XStringFormats.TopLeft);

            //Items
            XPen pen5 = new XPen(XColors.Gray, 0.5);
            pen5.DashStyle = XDashStyle.Solid;
            XGraphicsPath path5 = new XGraphicsPath();
            path5.AddLine(offsetX, 245, 570, 245);
            gfx.DrawPath(pen5, path5);

            rect = new XRect(offsetX + 1, 248, 540, 242);
            tf.DrawString("Line", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);
            //tf.DrawString("Item No.", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 40, 248, 540, 242);
            tf.DrawString("Item No.", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 125, 248, 540, 242);
            tf.DrawString("UOM", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 155, 248, 540, 242);
            tf.DrawString("Description", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 380, 248, 540, 242);
            tf.DrawString("Quantity", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 440, 248, 540, 242);
            tf.DrawString("Unit", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            rect = new XRect(offsetX + 495, 248, 540, 242);
            tf.DrawString("Extended", italicFont, XBrushes.Black, rect, XStringFormats.TopLeft);

            XPen pen6 = new XPen(XColors.Gray, 0.5);
            pen6.DashStyle = XDashStyle.Solid;
            XGraphicsPath path6 = new XGraphicsPath();
            path6.AddLine(offsetX, 258, 570, 258);
            gfx.DrawPath(pen6, path6);
            return gfx;
        }

        public string ReturnShortenedShippingDescription(string val)
        {
            switch (val)
            {
                case "PREMIUM DELIVERY FEE": val = "Prem. Delivery Fee"; break;
                case "INSTRUMENT HANDLING": val = "Instrument Handling"; break;
                case "HANDLING CHARGE": val = "Handling Charge"; break;
                case "FREIGHT CHARGE PRIMER": val = "Freight"; break;
                case "MIN HANDLING CHARGE FOR PRIMER": val = "Primer Handling"; break;
                case "INSTRUMENT FREIGHT 1": val = "Instrument Freight"; break;
                case "MIN HANDLING CHARGE FR CATALOG": val = "Min. Handling Charge"; break;
                case "SHIPPING AND HANDLING CHARGE": val = "Ship.&Handling"; break;
                case "DRY / WET ICE CHARGES": val = "Dry/Wet Ice Charge"; break;
                case "HAZARDOUS MATERIAL CHARGE": val = "HazMat Charge"; break;
                case "INSTRUMENT FREIGHT": val = "Instrument Freight"; break;
                case "DRY ICE CHARGE": val = "Dry Ice Charge"; break;
                case "RESTOCKING FEES": val = "Restock Fee"; break;
                case "CYLINDER RENT": val = "Rental Fee"; break;
                case "UZZZDEMANDCHGCYL": val = "Surcharge"; break;
                default: val = val.Length > 20 ? val.ToLower().Substring(20) : val; break;
            }

            if (val.Length == 0)
                val = "Misc.";

            return val;

        }
    }
}
