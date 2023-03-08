using Read_cXML_Invoices.Objects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Classes.Process_Invoice
{
    public class PrdAutoPostInvoice : AutoPostInvoice
    {
        public override void AutoPost(Ship ship, InvoiceHeader invoice)
        {
            try
            {
                ValidateInvoicePurchaseOrder(ship, invoice);
                if (invoice.Errors.Count == 0) UpdatePurchaseOrderAfter(invoice);
                if (invoice.Errors.Count == 0) InsertGlLines(invoice);
                if (invoice.Errors.Count == 0) UpdatePurchaseOrder(invoice);
                if (invoice.Errors.Count == 0 && invoice.PDFFileName.Length > 0) TagKwiktagDocument(ship, invoice);
                if (invoice.Errors.Count == 0 && invoice.PurchaseOrderNo.Length > 0) PostInvoice(ship, invoice);
            }
            catch (Exception ex)
            {
                invoice.Errors.Add(ex.Message);
                Constants.ERRORS.Add(new Error(ex, new System.Data.SqlClient.SqlCommand($"Invoice ID = '{invoice.InvoiceID}', Order ID = '{invoice.OrderID}'"), "AutoPostInvoice_PrdAutoPost", "AutoPost"));
            }

            UpdatePurchaseOrderAfter(invoice);
        }

        private void ValidateInvoicePurchaseOrder(Ship ship, InvoiceHeader invoice)
        {
            if (invoice.Lines.Count == 0)
                invoice.Errors.Add("Invoice does not have any lines!");
            else if (invoice.Document_ID.Length == 0)
            {
                invoice.Errors.Add("Purchase Order for Invoice could not be found in Navision.");
                invoice.PO_NAV_Status = "PO_NOT_FOUND";
            }
            else
            {
                PrdAutoPostDoc.AutoPostDocument auto = new PrdAutoPostDoc.AutoPostDocument();
                auto.UseDefaultCredentials = true;

                PrdPurchaseOrder.PurchaseOrder_Service poservice = new PrdPurchaseOrder.PurchaseOrder_Service();
                poservice.UseDefaultCredentials = true;

                PrdPurchaseOrder.PurchaseOrder po = poservice.Read(invoice.Document_ID);

                if (po == null)
                    invoice.Errors.Add("Purchase Order for Invoice could not be found in Navision.");
                else
                {
                    decimal invoiceamount = 0.0m, poamount = 0.0m;
                    foreach (var line in invoice.Lines)
                    {
                        string othererror = "";

                        invoiceamount += (line.Quantity * line.UnitPrice);
                        line.GSS_Part_Number = Database.GetItemNumber(po.No, line.SupplierPartID);

                        PrdPurchaseOrder.Purchase_Order_Line pline = null;
                        if (line.GSS_Part_Number.Length > 0)
                        {
                            foreach (var pl in po.PurchLines)
                            {
                                if (line.GSS_Part_Number == pl.No && invoice.Lines.Find(l => l.PurchLine_LineNumber == pl.Line_No) == null && line.Quantity <= (pl.Quantity - pl.Quantity_Invoiced))
                                {
                                    decimal qtyopen = pl.Quantity - pl.Quantity_Received;
                                    if (ship.ShipType == "NONDROP_SHIP" && line.Quantity <= pl.Quantity_Received)
                                        pline = pl;
                                    else if (ship.ShipType == "NONDROP_SHIP" && line.Quantity > pl.Quantity_Received)
                                        othererror = $"Inv. Line Qty of {line.Quantity.ToString("G29")} is greater than PO Line Qty Received of {pl.Quantity_Received.ToString("G29")}. Item has not been received yet.";
                                    else if (ship.ShipType == "NONDROP_SHIP" && line.Quantity >= qtyopen)
                                        othererror = $"Inv. Line Qty of {line.Quantity.ToString("G29")} is greater than/equal to PO Line Open Qty of {qtyopen.ToString("G29")}. Item has not been received yet.";
                                    else if (line.Quantity > 0.0M && (line.Quantity * line.UnitPrice) == 0.00M && (pl.Quantity * pl.Unit_Cost_LCY) > 0.00M)
                                        othererror = $"Inv. Line Amount is $0 and the PO Line Amount is ${(pl.Quantity * pl.Unit_Cost_LCY).ToString("G29")}";
                                    else if (ship.ShipType == "DROP_SHIP")
                                        pline = pl;
                                    break;
                                }
                                else if (line.GSS_Part_Number == pl.No && invoice.Lines.Find(l => l.PurchLine_LineNumber == pl.Line_No) == null && line.Quantity > (pl.Quantity - pl.Quantity_Invoiced))
                                    othererror = $"Inv. Line Qty of {line.Quantity.ToString("G29")} is greater than PO Line Qty Left of {(pl.Quantity - pl.Quantity_Invoiced).ToString("G29")}";
                            }
                        }

                        if (pline == null)
                            invoice.Errors.Add($"Invoice Line {line.LineNumber}, {line.SupplierPartID} could not be found in the Purchase Order {po.No} in Navision::{othererror}");
                        else
                        {
                            line.PurchLine_LineNumber = pline.Line_No;
                            poamount += pline.Quantity * pline.Unit_Cost_LCY;
                        }
                    }

                    if (invoice.Errors.Count == 0)
                    {
                        if (invoiceamount > (poamount + Constants.InvoiceThresholdAmt))
                            invoice.Errors.Add($"PRICE DISCREPANCY::Invoice Line Amount: ${invoiceamount.ToString("G29")}; PO Amount: ${poamount.ToString("G29")}");
                        else if (invoiceamount <= 0.00M)
                            invoice.Errors.Add($"Invoice is ${invoiceamount}.");

                        foreach (var line in invoice.ShipLines)
                        {
                            if ((line.Quantity * line.UnitPrice) > 0.0M)
                            {
                                line.GSS_Part_Number = Database.GetGlAccountNumber(ship.ShipType, line.SupplierPartID, line.Description);
                                if (line.GSS_Part_Number.Length == 0)
                                    invoice.Errors.Add($"Invoice GL Line {line.LineNumber}, {line.SupplierPartID}: {line.Description} GL Account Number could not be mapped and is empty!");
                            }
                        }

                        CheckInvoiceHeaderAmounts(ship, invoice);
                    }

                    if (invoice.Errors.Count == 0)
                    {
                        if (po.Status == PrdPurchaseOrder.Status.Released)
                        {
                            poservice.Update(ref po);
                            po = poservice.Read(po.No);

                            auto.ReopenPurchaseOrder(po.No);

                            po = poservice.Read(po.No);
                            poservice.Update(ref po);
                        }

                        invoice.PurchaseOrderNo = po.No;
                        invoice.BuyFromVendorNo = po.Buy_from_Vendor_No;
                        invoice.BuyFromVendorName = po.Buy_from_Vendor_Name;
                    }
                }
                auto.Dispose();
                poservice.Dispose();
            }
        }

        private void InsertGlLines(InvoiceHeader invoice)
        {
            if (invoice.ShipLines.Count > 0 && invoice.Vendor.ToUpper() == "EPPENDORF")
            {
                if (invoice.ShipType == "NONDROP_SHIP")
                {
                    invoice.Notes = $"GL Accounts not inserted for {invoice.Vendor} invoices.";
                    return;
                }
                else
                {

                }
            }
            else if (invoice.ShipLines.Count > 0 && (invoice.Vendor.ToUpper() == "GENEWIZ"))
            {
                invoice.Errors.Add($"GL Accounts are not allowed for {invoice.Vendor} invoices.");
                return;
            }

            decimal shipLineAmounts = 0.0m;

            foreach (var shipline in invoice.ShipLines)
            {
                if ((shipline.Quantity * shipline.UnitPrice) > 0.0M)
                {
                    try
                    {
                        PrdAutoPostDoc.AutoPostDocument auto = new PrdAutoPostDoc.AutoPostDocument();
                        auto.UseDefaultCredentials = true;

                        int lineNumber = auto.InsertPoGlLine(invoice.PurchaseOrderNo, shipline.GSS_Part_Number, shipline.Description, shipline.UnitPrice, shipline.Quantity);
                        if (lineNumber > 0)
                            shipline.PurchLine_LineNumber = lineNumber;
                        else
                            invoice.Errors.Add($"Adding GL Line {shipline.GSS_Part_Number} for the amount of {shipline.UnitPrice.ToString("C")} failed.");
                    }
                    catch (Exception ex)
                    {
                        Constants.ERRORS.Add(new Error(ex, null, "DevAutoPostInvoice", "InsertGlLines"));
                        invoice.Errors.Add(ex.Message);
                    }

                    shipLineAmounts += (shipline.UnitPrice * shipline.Quantity);
                }
            }

            if (shipLineAmounts > 250.00M)
                invoice.Errors.Add($"Invoice GL Amount Total is over $250.00!");
        }

        private void UpdatePurchaseOrder(InvoiceHeader invoice)
        {
            PrdAutoPostDoc.AutoPostDocument auto = new PrdAutoPostDoc.AutoPostDocument();
            auto.UseDefaultCredentials = true;

            PrdPurchaseOrder.PurchaseOrder_Service poservice = new PrdPurchaseOrder.PurchaseOrder_Service();
            poservice.UseDefaultCredentials = true;

            PrdPurchaseOrder.PurchaseOrder po = poservice.Read(invoice.PurchaseOrderNo);

            poservice.Update(ref po);
            po = poservice.Read(po.No);

            if (invoice.Vendor.ToUpper().Contains("VWR") && invoice.InvoiceDetailDiscount != 0)
            {
                po.Vendor_Invoice_No = invoice.InvoiceID + "-NET20";
                po.Vendor_Invoice_No_Confirm = invoice.InvoiceID + "-NET20";
            }
            else
            {
                po.Vendor_Invoice_No = invoice.InvoiceID;
                po.Vendor_Invoice_No_Confirm = invoice.InvoiceID;
            }

            po.Document_DateSpecified = true;
            DateTime invoiceDate = DateTime.Now;
            try { invoiceDate = DateTime.Parse(invoice.InvoiceDate); }
            catch { invoiceDate = invoice.ReceiveDate; }
            po.Document_Date = invoiceDate.Date;

            po.Posting_DateSpecified = true;
            po.Posting_Date = DateTime.Now.Date;
            po.dA_Batch_ID = invoice.dABatchID.ToString();

            poservice.Update(ref po);
            po = poservice.Read(po.No);

            foreach (var line in invoice.Lines)
            {
                PrdPurchaseOrder.Purchase_Order_Line pline = null;
                foreach (var pl in po.PurchLines)
                    if (pl.Type == PrdPurchaseOrder.Type.Item && pl.Line_No == line.PurchLine_LineNumber)
                        pline = pl;

                if (pline != null)
                {
                    if (line.Quantity <= (pline.Quantity - pline.Quantity_Received))
                    {
                        pline.docAlpha_Qty_to_ReceiveSpecified = true;
                        pline.docAlpha_Qty_to_Receive = line.Quantity;
                    }

                    pline.docAlpha_Qty_to_InvoiceSpecified = true;
                    pline.docAlpha_Qty_to_Invoice = line.Quantity;

                    pline.docAlpha_Direct_Unit_CostSpecified = true;
                    pline.docAlpha_Direct_Unit_Cost = line.UnitPrice;

                    poservice.Update(ref po);
                    po = poservice.Read(po.No);
                }
            }

            foreach (var line in invoice.ShipLines)
            {
                PrdPurchaseOrder.Purchase_Order_Line pline = null;
                foreach (var pl in po.PurchLines)
                    if (pl.Line_No == line.PurchLine_LineNumber)
                        pline = pl;

                if (pline != null)
                {
                    if (pline.docAlpha_Direct_Unit_Cost == 0.0M)
                    {
                        pline.docAlpha_Direct_Unit_CostSpecified = true;
                        pline.docAlpha_Direct_Unit_Cost = line.UnitPrice;
                    }

                    if (pline.docAlpha_Qty_to_Receive == 0)
                    {
                        pline.docAlpha_Qty_to_ReceiveSpecified = true;
                        pline.docAlpha_Qty_to_Receive = line.Quantity;
                    }

                    if (pline.docAlpha_Qty_to_Invoice == 0)
                    {
                        pline.docAlpha_Qty_to_InvoiceSpecified = true;
                        pline.docAlpha_Qty_to_Invoice = line.Quantity;
                    }

                    poservice.Update(ref po);
                    po = poservice.Read(po.No);
                }
            }

            List<int> lineNos = new List<int>();
            foreach (var line in invoice.Lines)
            {
                lineNos.Add(line.PurchLine_LineNumber);
                invoice.CalculatedInvoiceTotal += (line.Quantity * line.UnitPrice);
            }
            foreach (var line in invoice.ShipLines)
            {
                lineNos.Add(line.PurchLine_LineNumber);
                invoice.CalculatedInvoiceTotal += (line.Quantity * line.UnitPrice);
            }

            if (lineNos.Count > 0)
            {
                invoice.PurchaseOrder_LineNos = lineNos;

                foreach (var line in po.PurchLines)
                {
                    if (!lineNos.Contains(line.Line_No))
                    {
                        line.docAlpha_Qty_to_ReceiveSpecified = true;
                        line.docAlpha_Qty_to_Receive = 0;
                        line.docAlpha_Qty_to_InvoiceSpecified = true;
                        line.docAlpha_Qty_to_Invoice = 0;

                        po = poservice.Read(po.No);
                        poservice.Update(ref po);
                    }
                }
            }

            poservice.Update(ref po);
            po = poservice.Read(po.No);

            poservice.Dispose();
        }

        private void PostInvoice(Ship ship, InvoiceHeader invoice)
        {
            PrdAutoPostDoc.AutoPostDocument auto = new PrdAutoPostDoc.AutoPostDocument();
            auto.UseDefaultCredentials = true;
            string poNumber = invoice.PurchaseOrderNo;

            if (ship.ShipType == "DROP_SHIP")
            {
                int step = 0;
                try
                {
                    auto.PoPrepareDocAlpha(poNumber, true, false);
                    step++;

                    auto.AutoPostPo(poNumber, true, false);
                    step++;
                    invoice.PurchaseOrderPostedReceipt = true;
                }
                catch (Exception ex)
                {
                    string innermsg = "";
                    if (step == 0) innermsg = "Could not prepare PO";
                    else if (step == 1) innermsg = "Could not post PO";

                    invoice.Errors.Add($"Could not post receive::{innermsg}: {ex.Message}");
                    invoice.PurchaseOrderPostedReceipt = false;
                }

                //if (invoice.Errors.Count == 0 && invoice.ErrorCode.Length == 0)
                if (invoice.Errors.Count == 0)
                {
                    try
                    {
                        step = 0;

                        auto.PoPrepareDocAlpha(poNumber, false, true);
                        step++;

                        invoice.PurchaseOrderPostedTotal = Database.GetTotals(invoice.PurchaseOrderNo, invoice.PurchaseOrder_LineNos);

                        auto.AutoPostPo(poNumber, false, true);
                        step++;
                        invoice.PurchaseOrderPostedInvoice = true;
                    }
                    catch (Exception ex)
                    {
                        string innermsg = "";
                        if (step == 0) innermsg = "Could not prepare PO";
                        else if (step == 1) innermsg = "Could not post PO";

                        invoice.Errors.Add($"Could not post invoice::{innermsg}: {ex.Message}");
                        invoice.PurchaseOrderPostedInvoice = false;
                    }
                }
            }
            else //if (invoice.ErrorCode.Length == 0)
            {
                int step = 0;
                try
                {
                    auto.PoPrepareDocAlpha(poNumber, true, true);
                    step++;

                    invoice.PurchaseOrderPostedTotal = Database.GetTotals(invoice.PurchaseOrderNo, invoice.PurchaseOrder_LineNos);

                    auto.AutoPostPo(poNumber, true, true);
                    step++;

                    invoice.PurchaseOrderPostedReceipt = true;
                    invoice.PurchaseOrderPostedInvoice = true;
                }
                catch (Exception ex)
                {
                    string innermsg = "";
                    if (step == 0) innermsg = "Could not prepare PO";
                    else if (step == 1) innermsg = "Could not post PO";

                    invoice.Errors.Add($"Could not post receive and invoice::{innermsg}: {ex.Message}");
                    invoice.PurchaseOrderPostedReceipt = false;
                    invoice.PurchaseOrderPostedInvoice = false;
                }
            }

            auto.Dispose();
        }

        private void UpdatePurchaseOrderAfter(InvoiceHeader invoice)
        {
            try
            {
                PrdPurchaseOrder.PurchaseOrder_Service poservice = new PrdPurchaseOrder.PurchaseOrder_Service();
                poservice.UseDefaultCredentials = true;

                if (invoice.PurchaseOrderNo.Length > 0)
                {
                    PrdPurchaseOrder.PurchaseOrder po = poservice.Read(invoice.PurchaseOrderNo);

                    if (po != null)
                    {
                        if (poservice.IsUpdated(po.Key))
                            poservice.Read(po.Key);

                        foreach (var pline in po.PurchLines)
                        {
                            if (pline != null && (pline.Type == PrdPurchaseOrder.Type.Item || pline.Type == PrdPurchaseOrder.Type.G_L_Account))
                            {
                                pline.docAlpha_Qty_to_ReceiveSpecified = true;
                                pline.docAlpha_Qty_to_Receive = 0;

                                pline.docAlpha_Qty_to_InvoiceSpecified = true;
                                pline.docAlpha_Qty_to_Invoice = 0;

                                pline.docAlpha_Direct_Unit_CostSpecified = true;
                                pline.docAlpha_Direct_Unit_Cost = 0;

                                pline.Qty_to_InvoiceSpecified = true;
                                pline.Qty_to_Invoice = pline.Quantity - pline.Quantity_Invoiced;

                                pline.Quantity_ReceivedSpecified = true;
                                pline.Qty_to_Receive = pline.Quantity - pline.Quantity_Received;

                                poservice.Update(ref po);
                                po = poservice.Read(po.No);
                            }

                            if (pline != null && (pline.Type == PrdPurchaseOrder.Type.G_L_Account) && invoice.Errors.Count > 0)
                            {
                                pline.QuantitySpecified = true;
                                pline.Quantity = 0;

                                pline.Qty_to_InvoiceSpecified = true;
                                pline.Qty_to_Invoice = 0;

                                pline.Qty_to_ReceiveSpecified = true;
                                pline.Qty_to_Receive = 0;

                                poservice.Update(ref po);
                                po = poservice.Read(po.No);
                            }
                        }

                        poservice.Update(ref po);
                        po = poservice.Read(po.No);
                    }
                }
                else
                    throw new Exception($"Purchase Order Number for {invoice.Vendor} invoice # {invoice.InvoiceID} was empty!");
            }
            catch (Exception ex)
            {
                invoice.Errors.Add(ex.Message);
                Constants.ERRORS.Add(new Error(ex, new System.Data.SqlClient.SqlCommand($"Invoice ID = '{invoice.InvoiceID}', Order ID = '{invoice.OrderID}'"), "AutoPostInvoice_PrdAutoPost", "UpdatePurchaseOrderAfter"));
            }
        }
    }
}
