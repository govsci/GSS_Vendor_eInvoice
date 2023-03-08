using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Read_cXML_Invoices.Objects;
using System.IO;

namespace Read_cXML_Invoices.Classes.Report
{
    public class SendCsvReport
    {
        private List<CsvFiles> Files = null;

        public SendCsvReport(List<InvoiceHeader> invoices, List<Ship> ships, List<InvoiceHeader> invoicesOnHold, List<InvoiceHeader> emptyVendorInvoices, DateTime appStarted)
        {
            Files = new List<CsvFiles>();
            AddBatchInfoFile(ships);
            decimal batchTotal = AddInvoicesReceivedFile(invoices);

            invoicesOnHold = invoicesOnHold.OrderBy(i => i.Vendor).ThenBy(i => i.ReleaseDate).ToList();
            AddInvoicesOnHoldFile(invoicesOnHold);

            AddHandsOnInvoicesFile(ships);
            AddPoNotFoundFile(ships);
            AddPostedInvoicesFile(ships);
            AddPostedInvoicesFailedFile(ships);
            AddEmptyVendorInvoicesFile(emptyVendorInvoices);
            WriteToFile();
        }

        private void AddBatchInfoFile(List<Ship> ships)
        {
            StringBuilder csvContent = new StringBuilder();

            csvContent.AppendLine("Ship Type,Batch Number,Vendor,dABatchID,# of Invoices,Total Amount");

            foreach (Ship ship in ships)
                foreach (Batch batch in ship.Batches)
                    foreach (Vendor vendor in batch.Vendors)
                        csvContent.AppendLine($"{ship.ShipType},{batch.BatchNumber},{vendor.VendorName},{vendor.daBatchId},{vendor.Invoices.Count},{vendor.Total.ToString("G29").Replace(",", "")}");

            foreach (Ship ship in ships)
                foreach (Batch batch in ship.PoNotFoundBatches)
                    foreach (Vendor vendor in batch.Vendors)
                        csvContent.AppendLine($"{ship.ShipType},{batch.BatchNumber},{vendor.VendorName},{vendor.daBatchId},{vendor.Invoices.Count},{vendor.Total.ToString("G29").Replace(",", "")}");

            Files.Add(new CsvFiles("Batch Information", csvContent.ToString()));
        }
        private decimal AddInvoicesReceivedFile(List<InvoiceHeader> invoices)
        {
            decimal batchtotal = 0.0m;
            StringBuilder csvContent = new StringBuilder();

            csvContent.AppendLine("Vendor Name"
                + ",Invoice ID"
                + ",Invoice Date"
                + ",PO #"
                + ",PO Date"
                + ",Subtotal Amount"
                + ",Tax"
                + ",Shipping Amount"
                + ",Special Handling Amount"
                + ",Invoice Total"
                + ",Line Number"
                + ",Part Number"
                + ",Description"
                + ",Unit of Measure"
                + ",Quantity"
                + ",Unit Price"
                + ",Line Tax"
                + ",Line Total"
                + ",Receive Date");

            foreach (InvoiceHeader invoice in invoices)
            {
                foreach (InvoiceLine line in invoice.Lines)
                {
                    string invoiceTotal = "";
                    if (invoice.DueAmount > 0.0M)
                        invoiceTotal = invoice.DueAmount.ToString("G29");
                    else if (invoice.NetAmount > 0.00M)
                        invoiceTotal = invoice.NetAmount.ToString("G29");
                    else if (invoice.GrossAmount > 0.00M)
                        invoiceTotal = invoice.GrossAmount.ToString("G29");
                    invoiceTotal = invoiceTotal.Replace(",", "");

                    csvContent.AppendLine($"{invoice.Vendor}"
                        + $",{invoice.InvoiceID}"
                        + $",{invoice.InvoiceDate}"
                        + $",{invoice.OrderID}"
                        + $",{invoice.OrderDate}"
                        + $",{invoice.SubTotalAmount.ToString("G29").Replace(",", "")}"
                        + $",{invoice.Tax.ToString("G29").Replace(",", "")}"
                        + $",{invoice.ShippingAmount.ToString("G29").Replace(",", "")}"
                        + $",{invoice.SpecialHandlingAmount.ToString("G29").Replace(",", "")}"
                        + $",{invoiceTotal}"
                        + $",{line.LineNumber}"
                        + $",{line.SupplierPartID.Replace(",", "")}"
                        + $",{line.Description.Replace(",", "")}"
                        + $",{line.UnitOfMeasure}"
                        + $",{line.Quantity.ToString("G29").Replace(",", "")}"
                        + $",{line.UnitPrice.ToString("G29").Replace(",", "")}"
                        + $",{line.Tax.ToString("G29").Replace(",", "")}"
                        + $",{line.LineTotal.ToString("G29").Replace(",", "")}"
                        + $",{invoice.ReceiveDate.ToString("MM/dd/yyyy hh:mm tt")}");

                    batchtotal += (line.Quantity * line.UnitPrice);
                }
            }

            Files.Add(new CsvFiles("Invoices Received Report", csvContent.ToString()));

            return batchtotal;
        }
        private void AddHandsOnInvoicesFile(List<Ship> ships)
        {
            int count = 0;

            foreach (Ship ship in ships)
                foreach (Batch batch in ship.PoNotFoundBatches)
                    foreach (Vendor vendor in batch.Vendors)
                        foreach (InvoiceHeader invoice in vendor.Invoices)
                            if (!invoice.InvoiceRetry)
                                count++;

            if (count > 0)
            {
                StringBuilder csvContent = new StringBuilder();

                csvContent.AppendLine("Ship Type"
                    + ",Batch No."
                    + ",Vendor Name"
                    + ",Invoice ID"
                    + ",Invoice Date"
                    + ",Received Date"
                    + ",PO #"
                    + ",PO Date"
                    + ",Subtotal Amount"
                    + ",Tax"
                    + ",Shipping"
                    + ",Special Handling"
                    + ",Invoice Total"
                    + ",Status Code");

                foreach (Ship ship in ships)
                {
                    foreach (Batch batch in ship.PoNotFoundBatches)
                    {
                        foreach (Vendor vendor in batch.Vendors)
                        {
                            foreach (InvoiceHeader invoice in vendor.Invoices)
                            {
                                if (!invoice.InvoiceRetry && invoice.PO_NAV_Status != "PO_POSTED")
                                {
                                    string invoiceTotal = "";
                                    if (invoice.DueAmount > 0.0M)
                                        invoiceTotal = invoice.DueAmount.ToString("G29");
                                    else if (invoice.NetAmount > 0.00M)
                                        invoiceTotal = invoice.NetAmount.ToString("G29");
                                    else if (invoice.GrossAmount > 0.00M)
                                        invoiceTotal = invoice.GrossAmount.ToString("G29");
                                    invoiceTotal = invoiceTotal.Replace(",", "");

                                    csvContent.AppendLine($"{ship.ShipType}"
                                        + $",{batch.BatchNumber}"
                                        + $",{vendor.VendorName}"
                                        + $",{invoice.InvoiceID}"
                                        + $",{invoice.InvoiceDate}"
                                        + $",{invoice.ReceiveDate.ToShortDateString()}"
                                        + $",{invoice.OrderID}"
                                        + $",{invoice.OrderDate}"
                                        + $",{invoice.SubTotalAmount.ToString("G29").Replace(",", "")}"
                                        + $",{invoice.Tax.ToString("G29").Replace(",", "")}"
                                        + $",{invoice.ShippingAmount.ToString("G29").Replace(",", "")}"
                                        + $",{invoice.SpecialHandlingAmount.ToString("G29").Replace(",", "")}"
                                        + $",{invoiceTotal}"
                                        + $",{invoice.PO_NAV_Status}");
                                }
                            }
                        }
                    }
                }

                Files.Add(new CsvFiles("Hands On Invoices", csvContent.ToString()));
            }
        }
        private void AddPoNotFoundFile(List<Ship> ships)
        {
            StringBuilder csvContent = new StringBuilder();

            csvContent.AppendLine(",Ship Type"
                + ",Batch No."
                + ",Vendor Name"
                + ",Invoice ID"
                + ",Invoice Date"
                + ",Received Date"
                + ",PO #"
                + ",PO Date"
                + ",Subtotal Amount"
                + ",Tax"
                + ",Shipping"
                + ",Special Handling"
                + ",Invoice Total"
                + ",Retry?"
                + ",Status Code");

            foreach (Ship ship in ships)
            {
                foreach (Batch batch in ship.PoNotFoundBatches)
                {
                    foreach (Vendor vendor in batch.Vendors)
                    {
                        foreach (InvoiceHeader invoice in vendor.Invoices)
                        {
                            string invoiceTotal = "";
                            if (invoice.DueAmount > 0.0M)
                                invoiceTotal = invoice.DueAmount.ToString("G29");
                            else if (invoice.NetAmount > 0.00M)
                                invoiceTotal = invoice.NetAmount.ToString("G29");
                            else if (invoice.GrossAmount > 0.00M)
                                invoiceTotal = invoice.GrossAmount.ToString("G29");
                            invoiceTotal = invoiceTotal.Replace(",", "");

                            csvContent.AppendLine($"{ship.ShipType}"
                                + $",{batch.BatchNumber}"
                                + $",{vendor.VendorName}"
                                + $",{invoice.InvoiceID}"
                                + $",{invoice.InvoiceDate}"
                                + $",{invoice.ReceiveDate.ToShortDateString()}"
                                + $",{invoice.OrderID}"
                                + $",{invoice.OrderDate}"
                                + $",{invoice.SubTotalAmount.ToString("G29").Replace(",", "")}"
                                + $",{invoice.Tax.ToString("G29").Replace(",", "")}"
                                + $",{invoice.ShippingAmount.ToString("G29").Replace(",", "")}"
                                + $",{invoice.SpecialHandlingAmount.ToString("G29").Replace(",", "")}"
                                + $",{invoiceTotal}"
                                + $",{(invoice.InvoiceRetry ? "Yes" : "No")}"
                                + $",{invoice.PO_NAV_Status}");
                        }
                    }
                }
            }
            Files.Add(new CsvFiles("PO Not Found Invoices", csvContent.ToString()));
        }
        private void AddInvoicesOnHoldFile(List<InvoiceHeader> invoices)
        {
            StringBuilder csvContents = new StringBuilder();

            csvContents.AppendLine("Vendor Name"
                + ",Invoice ID"
                + ",Invoice Date"
                + ",PO #"
                + ",PO Date"
                + ",Invoice Total"
                + ",Date to Release to DocAlpha");

            foreach (InvoiceHeader invoice in invoices)
            {
                string invoiceTotal = "";
                if (invoice.DueAmount > 0.0M)
                    invoiceTotal = invoice.DueAmount.ToString("G29");
                else if (invoice.NetAmount > 0.00M)
                    invoiceTotal = invoice.NetAmount.ToString("G29");
                else if (invoice.GrossAmount > 0.00M)
                    invoiceTotal = invoice.GrossAmount.ToString("G29");
                invoiceTotal = invoiceTotal.Replace(",", "");

                DateTime dateChecker;
                try { dateChecker = DateTime.Parse(invoice.InvoiceDate); }
                catch { dateChecker = invoice.ReceiveDate; }

                csvContents.AppendLine($"{invoice.Vendor}"
                    + $",{invoice.InvoiceID}"
                    + $",{dateChecker.ToShortDateString()}"
                    + $",{invoice.OrderID}"
                    + $",{invoice.OrderDate}"
                    + $",{invoiceTotal}"
                    + $",{invoice.ReleaseDate.ToShortDateString()}");
            }
            
            Files.Add(new CsvFiles("NDS Invoices On Hold", csvContents.ToString()));
        }
        private void AddPostedInvoicesFile(List<Ship> ships)
        {
            StringBuilder csvContent = new StringBuilder();

            csvContent.AppendLine("Ship Type"
                + ",Batch No."
                + ",Vendor Name"
                + ",dA Batch ID"
                + ",Invoice ID"
                + ",Invoice Date"
                + ",Received Date"
                + ",PO #"
                + ",PO Date"
                + ",Subtotal Amount"
                + ",Tax"
                + ",GL Accounts Amount"
                + ",Invoice Total"
                + ",Purchase Order Posted Total"
                + ",Totals Match?"
                + ",PO Posted Receipt"
                + ",PO Posted Invoice"
                + ",Kwiktagged?"
                + ",Notes");
                

            foreach (var ship in ships)
            {
                foreach (var batch in ship.Batches)
                {
                    foreach (var vendor in batch.Vendors)
                    {
                        foreach (var invoice in vendor.Invoices)
                        {
                            if (invoice.PurchaseOrderPostedInvoice && invoice.PurchaseOrderPostedReceipt && invoice.Errors.Count == 0)
                            {
                                decimal specialAmountTotal = 0.0M;
                                foreach (var shipLine in invoice.ShipLines)
                                    specialAmountTotal += (shipLine.Quantity * shipLine.UnitPrice);

                                csvContent.AppendLine($"{ship.ShipType}"
                                    + $",{batch.BatchNumber.ToString()}"
                                    + $",{vendor.VendorName}"
                                    + $",{vendor.daBatchId}"
                                    + $",{invoice.InvoiceID}"
                                    + $",{invoice.InvoiceDate}"
                                    + $",{invoice.ReceiveDate.ToShortDateString()}"
                                    + $",{invoice.OrderID}"
                                    + $",{invoice.OrderDate}"
                                    + $",{invoice.SubTotalAmount.ToString("G29").Replace(",", "")}"
                                    + $",{invoice.Tax.ToString("G29").Replace(",", "")}"
                                    + $",{specialAmountTotal.ToString("G29").Replace(",", "")}"
                                    + $",{invoice.CalculatedInvoiceTotal.ToString("G29").Replace(",", "")}"
                                    + $",{invoice.PurchaseOrderPostedTotal.ToString("G29").Replace(",", "")}"
                                    + $",{((invoice.CalculatedInvoiceTotal == invoice.PurchaseOrderPostedTotal) ? "Yes" : "No")}"
                                    + $",{(invoice.PurchaseOrderPostedReceipt ? "Yes" : "No")}"
                                    + $",{(invoice.PurchaseOrderPostedInvoice ? "Yes" : "No")}"
                                    + $",{(invoice.Kwiktagged ? "Yes" : "No")}"
                                    + $",{invoice.Notes.Replace(",", "")}");
                            }
                        }
                    }
                }
            }

            Files.Add(new CsvFiles("Posted Purchase Orders", csvContent.ToString()));
        }
        private void AddPostedInvoicesFailedFile(List<Ship> ships)
        {
            StringBuilder csvContent = new StringBuilder();

            csvContent.AppendLine("Ship Type"
                + ",Batch No."
                + ",Vendor Name"
                + ",Invoice ID"
                + ",Invoice Date"
                + ",Received Date"
                + ",PO #"
                + ",PO Date"
                + ",Subtotal Amount"
                + ",Tax"
                + ",GL Accounts Amount"
                + ",Invoice Total"
                + ",Purchase Order Posted Total"
                + ",Totals Match?"
                + ",PO Posted Receipt"
                + ",PO Posted Invoice"
                + ",Kwiktagged"
                + ",Errors");

            foreach (var ship in ships)
            {
                foreach (var batch in ship.Batches)
                {
                    foreach (var vendor in batch.Vendors)
                    {
                        foreach (var invoice in vendor.Invoices)
                        {
                            if (invoice.Errors.Count > 0)
                            {
                                decimal specialAmountTotal = 0.0M;
                                foreach (var shipLine in invoice.ShipLines)
                                    specialAmountTotal += (shipLine.Quantity * shipLine.UnitPrice);

                                string errors = "";
                                foreach (string error in invoice.Errors)
                                    errors += error + Environment.NewLine;

                                csvContent.AppendLine($"{ship.ShipType}"
                                    + $",{batch.BatchNumber.ToString()}"
                                    + $",{invoice.Vendor}"
                                    + $",{invoice.InvoiceID}"
                                    + $",{invoice.InvoiceDate}"
                                    + $",{invoice.ReceiveDate.ToShortDateString()}"
                                    + $",{invoice.OrderID}"
                                    + $",{invoice.OrderDate}"
                                    + $",{invoice.SubTotalAmount.ToString("G29").Replace(",","")}"
                                    + $",{invoice.Tax.ToString("G29").Replace(",", "")}"
                                    + $",{specialAmountTotal.ToString("G29").Replace(",", "")}"
                                    + $",{invoice.CalculatedInvoiceTotal.ToString("G29").Replace(",", "")}"
                                    + $",{invoice.PurchaseOrderPostedTotal.ToString("G29").Replace(",", "")}"
                                    + $",{((invoice.CalculatedInvoiceTotal == invoice.PurchaseOrderPostedTotal) ? "Yes" : "No")}"
                                    + $",{(invoice.PurchaseOrderPostedReceipt ? "Yes" : "No")}"
                                    + $",{(invoice.PurchaseOrderPostedInvoice ? "Yes" : "No")}"
                                    + $",{(invoice.Kwiktagged ? "Yes" : "No")}"
                                    + $",{errors.Replace(",", "")}");
                            }
                        }
                    }
                }
            }

            Files.Add(new CsvFiles("Invoices With Errors", csvContent.ToString()));
        }
        private void AddEmptyVendorInvoicesFile(List<InvoiceHeader> invoices)
        {
            StringBuilder csvContent = new StringBuilder();

            csvContent.AppendLine("Vendor Name"
                + ",Invoice ID"
                + ",Invoice Date"
                + ",PO #"
                + ",PO Date"
                + ",Subtotal Amount"
                + ",Tax"
                + ",Shipping Amount"
                + ",Special Handling Amount"
                + ",Invoice Total"
                + ",Receive Date");

            foreach (InvoiceHeader invoice in invoices)
            {
                string invoiceTotal = "";
                if (invoice.DueAmount > 0.0M)
                    invoiceTotal = invoice.DueAmount.ToString("G29");
                else if (invoice.NetAmount > 0.00M)
                    invoiceTotal = invoice.NetAmount.ToString("G29");
                else if (invoice.GrossAmount > 0.00M)
                    invoiceTotal = invoice.GrossAmount.ToString("G29");
                invoiceTotal = invoiceTotal.Replace(",", "");

                csvContent.AppendLine($"{invoice.Vendor}"
                    + $",{invoice.InvoiceID}"
                    + $",{invoice.InvoiceDate}"
                    + $",{invoice.OrderID}"
                    + $",{invoice.OrderDate}"
                    + $",{invoice.SubTotalAmount.ToString("G29").Replace(",", "")}"
                    + $",{invoice.Tax.ToString("G29").Replace(",", "")}"
                    + $",{invoice.ShippingAmount.ToString("G29").Replace(",", "")}"
                    + $",{invoice.SpecialHandlingAmount.ToString("G29").Replace(",", "")}"
                    + $",{invoiceTotal}"
                    + $",{invoice.ReceiveDate.ToString("MM/dd/yyyy hh:mm tt")}");
            }

            Files.Add(new CsvFiles("Empty Vendor Invoices", csvContent.ToString()));
        }

        private void WriteToFile()
        {
            string path = Constants.ReportPath + DateTime.Now.ToString(@"yyyy\\MM\\dd\\");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            foreach (CsvFiles file in Files)
            {
                string fileName = $"{path}{DateTime.Now.ToString("yyyyMMddHHmmssffff")}.{file.SheetName}.csv";
                File.WriteAllText(fileName, file.Content);
            }
        }
    }
}
