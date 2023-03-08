using Read_cXML_Invoices.Objects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Classes.Process_Invoice
{
    public abstract class AutoPostInvoice
    {
        public abstract void AutoPost(Ship ship, InvoiceHeader invoice);

        public void TagKwiktagDocument(Ship ship, InvoiceHeader invoice)
        {
            try
            {
                /// // SET Variables
                string ktserverurl = Constants.KwiktagURL;
                string callingid = "877c0613-2b23-4647-b08d-0fb1fa4c35a0";
                string username = CustomTextEncrypt.Decode(Constants.KwiktagUserName);
                string password = CustomTextEncrypt.Decode(Constants.KwiktagPassword);

                // Load document into a byte array
                FileStream stream = File.OpenRead(invoice.PDFFileName);
                byte[] fileBytes = new byte[stream.Length];
                stream.Read(fileBytes, 0, fileBytes.Length);
                stream.Close();

                // Authenticate to KwikTag
                KwikTagSDKLibrary.Authentication oAuth = new KwikTagSDKLibrary.Authentication(ktserverurl, callingid);
                KwikTagSDKLibrary.XmlReturns.ConfigReturn oConfig = oAuth.AuthenticateKwikTagUserAccount(username, password);

                // Get A System Barocde
                KwikTagSDKLibrary.Barcode oBarcode = new KwikTagSDKLibrary.Barcode(ktserverurl, callingid, oConfig.Token, username);
                KwikTagSDKLibrary.XmlReturns.XferDataReturn oBarcodeReturn = new KwikTagSDKLibrary.XmlReturns.XferDataReturn();
                oBarcodeReturn = (KwikTagSDKLibrary.XmlReturns.XferDataReturn)KwikTagSDKLibrary.Utilities.XmlUtil.DeserializeFromXml(oBarcode.RetrieveSystem(1001), typeof(KwikTagSDKLibrary.XmlReturns.XferDataReturn));  // This call returns just a string value with the barcode

                // Create the Document Metadata container
                KwikTagSDKLibrary.Document oDoc = new KwikTagSDKLibrary.Document(ktserverurl, callingid, oConfig.Token, username);
                KwikTagSDKLibrary.KwikTagDocumentModel oKtDocument = new KwikTagSDKLibrary.KwikTagDocumentModel();

                // Load the Metadata
                KwikTagSDKLibrary.TagField oKtDocumentTagField = new KwikTagSDKLibrary.TagField();

                oKtDocumentTagField.Name = "Number";
                oKtDocumentTagField.TagValue = invoice.PurchaseOrderNo; //Purchase Header "No."
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "Buy From Vendor Name";
                oKtDocumentTagField.TagValue = invoice.BuyFromVendorName; //Purchase Header "Buy-from Vendor Name"
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "Vendor ID";
                oKtDocumentTagField.TagValue = invoice.BuyFromVendorNo; //Purchase Header "Buy-from Vendor No."
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "Document Date";
                oKtDocumentTagField.TagValue = DateTime.Now.ToString("yyyy-MM-dd"); //"2019-08-12";//DateTime.Now.ToString("yyyy-MM-dd");
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "Posting Description";
                oKtDocumentTagField.TagValue = "Order " + invoice.PurchaseOrderNo;
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "Company ID";
                oKtDocumentTagField.TagValue = "Government Scientific Source";
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "File Name";
                oKtDocumentTagField.TagValue = "";
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "Comments";
                oKtDocumentTagField.TagValue = ship.ShipType;
                oKtDocument.TagList.Add(oKtDocumentTagField);

                oKtDocumentTagField = new KwikTagSDKLibrary.TagField();
                oKtDocumentTagField.Name = "Quote Number";
                oKtDocumentTagField.TagValue = "";
                oKtDocument.TagList.Add(oKtDocumentTagField);

                // Set the document barcode, target drawer and target site values
                oKtDocument.Barcode = oBarcodeReturn.Data;
                oKtDocument.Drawer = "Purchase Order";
                oKtDocument.SiteName = "Government Scientific Source";

                // Pass the values to the API
                KwikTagSDKLibrary.XmlReturns.XferDataReturn oXferDataResult = new KwikTagSDKLibrary.XmlReturns.XferDataReturn();
                oXferDataResult = oDoc.UploadAndTag(oKtDocument, System.IO.Path.GetFileName(invoice.PDFFileName), fileBytes);

                // Output the result
                //txtResult.Text = oXferDataResult.Success.ToString();
                //txtMessage.Text = oXferDataResult.ErrorMessage.ToString();

                if (oXferDataResult.ErrorMessage.ToString().Length > 0)
                    throw new Exception(oXferDataResult.ErrorMessage);

                invoice.Kwiktagged = true;
            }
            catch (Exception ex)
            {
                invoice.Errors.Add("Kwiktag Failed: " + ex.ToString());
            }
        }

        public void CheckInvoiceHeaderAmounts(Ship ship, InvoiceHeader invoice)
        {
            decimal specialAmount = invoice.SpecialHandlingAmount, shipLineTotal = 0.0m;
            foreach (var shipline in invoice.ShipLines) shipLineTotal += (shipline.Quantity * shipline.UnitPrice);
            if (shipLineTotal == specialAmount) specialAmount = 0.0m;
            if (specialAmount > 0.00M)
            {
                string glaccountno = Database.GetGlAccountNumber(ship.ShipType, "Special Handling", "Special Handling");
                InvoiceLine line = new InvoiceLine();
                line.LineNumber = 0;
                line.Quantity = 1;
                line.UnitOfMeasure = "EA";
                line.UnitPrice = specialAmount;
                line.ReferenceLineNumber = 0;
                line.GSS_Part_Number = glaccountno;
                line.Description = "Special Handling";
                line.Tax = 0;
                line.LineTotal = specialAmount;
                line.PurchLine_LineNumber = 0;
                line.ShipLine = 1;
                invoice.ShipLines.Add(line);
            }

            if (invoice.ShippingAmount > 0.0M)
            {
                string glaccountno = Database.GetGlAccountNumber(ship.ShipType, "Shipping", "Shipping");
                if (invoice.ShipLines.Find(s => s.GSS_Part_Number == glaccountno && s.UnitPrice == invoice.ShippingAmount) == null)
                {
                    InvoiceLine line = new InvoiceLine();
                    line.LineNumber = 0;
                    line.Quantity = 1;
                    line.UnitOfMeasure = "EA";
                    line.UnitPrice = invoice.ShippingAmount;
                    line.ReferenceLineNumber = 0;
                    line.GSS_Part_Number = glaccountno;
                    line.Description = "Shipping";
                    line.Tax = 0;
                    line.LineTotal = invoice.ShippingAmount;
                    line.PurchLine_LineNumber = 0;
                    line.ShipLine = 1;
                    invoice.ShipLines.Add(line);
                }
            }

            if (invoice.InvoiceDetailDiscount != 0.0M)
            {
                decimal discount = invoice.InvoiceDetailDiscount;
                if (discount > 0.0M) discount = discount * -1.00M;

                string glaccountno = Database.GetGlAccountNumber(ship.ShipType, "Discount", "Discount");
                InvoiceLine line = new InvoiceLine();
                line.LineNumber = 0;
                line.Quantity = 1;
                line.UnitOfMeasure = "EA";
                line.UnitPrice = discount;
                line.ReferenceLineNumber = 0;
                line.GSS_Part_Number = glaccountno;
                line.Description = "Discount";
                line.Tax = 0;
                line.LineTotal = discount;
                line.PurchLine_LineNumber = 0;
                line.ShipLine = 1;
                invoice.ShipLines.Add(line);
            }
        }
    }
}
