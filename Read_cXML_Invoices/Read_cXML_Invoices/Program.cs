using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Read_cXML_Invoices.Classes;
using Read_cXML_Invoices.Classes.Process_Invoice;
using Read_cXML_Invoices.Classes.Report;
using System.Data;
using System.Data.SqlClient;
using Read_cXML_Invoices.Objects;


namespace Read_cXML_Invoices
{
    class Program
    {
        public static List<InvoiceHeader> invoices = new List<InvoiceHeader>();
        public static List<InvoiceHeader> invoicesOnHold = new List<InvoiceHeader>();
        public static List<InvoiceHeader> emptyVendorInvoices = new List<InvoiceHeader>();
        public static List<Ship> ships = new List<Ship>();
        public static bool ReportSent = false;
        public static bool ReportNeedstoBeSent = true;
        public DateTime AppStarted;

        public Program(DateTime date, DateTime appStarted)
        {
            AppStarted = appStarted;
            //TestFileFunction();

            if (Constants.ERRORS.Count == 0)
            {
                Constants.InvoiceDropFolder = CheckDirectory(Constants.InvoiceDropFolder, 0);
                CheckDirectory(Constants.InvoiceDropFolder + Constants.MasterBatchFolder);
                InvoiceChecker(date);
                Database.PullInvoices(ref invoices, ref invoicesOnHold, ref emptyVendorInvoices);
                if (invoices.Count > 0 || invoicesOnHold.Count > 0)
                {
                    ships = BuildSendInvoices();
                    SendReport report = new SendReport(invoices, ships, invoicesOnHold, emptyVendorInvoices, AppStarted);
                    if (report.ReportErrors.Count > 0)
                    {
                        Constants.ERRORS.AddRange(report.ReportErrors);
                        new SendCsvReport(invoices, ships, invoicesOnHold, emptyVendorInvoices, AppStarted);
                    }
                    ReportSent = true;
                }
                else
                    Constants.ERRORS.Add(new Error(new Exception("Invoices are NULL!?"), "Program", "Program"));

                new Send_Batch_Report(Constants.InvoiceDropFolder);
            }
            else
                ReportNeedstoBeSent = false;
        }

        private void TestFileFunction()
        {
            try
            {
                string folder = $"{Constants.InvoiceDropFolder}test\\";
                string file = "test.txt";
                string contents = "DOES IT WORK";

                if (folder.EndsWith("\\test\\"))
                {
                    Directory.CreateDirectory(folder);
                    File.WriteAllText(folder + file, contents);

                    string readcontents = File.ReadAllText(folder + file);
                    if (readcontents != contents)
                        throw new Exception("Writing and reading the test file failed. Read_cXML_Invoices will not run.");
                    else
                        Directory.Delete(folder, true);
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, "Program", "TestFileFunction"));
            }
        }

        private void InvoiceChecker(DateTime date)
        {
            for (DateTime time = date; time.Date <= DateTime.Today.Date; time = time.AddDays(1))
            {
                Console.WriteLine(Constants.InvoiceFolder + time.ToString("yyyy") + "\\" + time.ToString("MM") + "\\" + time.ToString("dd"));
                CheckInvoiceFolder(Constants.InvoiceFolder + time.ToString("yyyy") + "\\" + time.ToString("MM") + "\\" + time.ToString("dd"));
            }
        }
        private void CheckInvoiceFolder(string path)
        {
            if (Directory.Exists(path))
            {
                foreach (string file in Directory.GetFiles(path))
                {
                    ReadXML read = new ReadXML(file);
                    if (read.Invoice != null)
                    {
                        InvoiceHeader invoice = read.Invoice;
                        UploadData upload = new UploadData(invoice);
                        upload.Upload();
                        Console.WriteLine(file);
                    }
                }
            }
        }

        private List<Ship> BuildSendInvoices()
        {
            List<Ship> ships = new List<Ship>();

            if (invoices.Count > 0)
            {
                foreach (InvoiceHeader invoice in invoices)
                {
                    Ship ship = ships.Find(s => s.ShipType == invoice.ShipType);
                    if (ship == null)
                    {
                        ship = new Ship(invoice.ShipType);
                        ships.Add(ship);
                    }

                    if (!invoice.PO_Found)
                        ship.PoNotFoundBatches = CheckBatchList(ship.PoNotFoundBatches, 1, invoice);
                    else
                        ship.Batches = CheckBatchList(ship.Batches, 1, invoice);
                }

                foreach (Ship ship in ships)
                {
                    foreach (Batch batch in ship.Batches)
                    {
                        foreach (Vendor vendor in batch.Vendors)
                        {
                            int dabatchid = Database.InsertBatchInformation(ship.ShipType, 0, Constants.BatchFolder + batch.BatchNumber, vendor.VendorName);
                            foreach (InvoiceHeader invoice in vendor.Invoices)
                            {
                                string dir = "";
                                try
                                {
                                    vendor.daBatchId = Constants.daBatchIDPreq + dabatchid.ToString();
                                    invoice.dABatchID = vendor.daBatchId;
                                    dir = Constants.InvoiceDropFolder + @"\" + ship.ShipType + @"\" + Constants.BatchFolder + batch.BatchNumber + @"\" + vendor.VendorName + @"\";
                                    CheckDirectory(dir);
                                    if (invoice.PDFFileName.Length == 0 || !File.Exists(invoice.PDFFileName))
                                    {
                                        BuildSingleInvoicePDF b = new BuildSingleInvoicePDF(invoice, dir, Constants.InvoiceDropFolder + @"\" + Constants.MasterBatchFolder);
                                        invoice.PDFFileName = b.CreatePDF();
                                        if (invoice.PDFFileName.Length == 0)
                                            invoice.Errors.Add("PDF File for invoice is empty!");
                                        //Email.SendEmail("Please see attached", invoice.Vendor + " eInvoice " + invoice.InvoiceID, "", "invoicing@govsci.com", "", "", invoice.PDFFileName, true);                                        
                                    }
                                    
                                    AutoPostInvoice auto;
                                    switch (Constants.AppProfile)
                                    {
                                        case "dev":
                                            auto = new DevAutoPostInvoice();
                                            auto.AutoPost(ship, invoice);
                                            break;
                                        case "prd":
                                            auto = new PrdAutoPostInvoice();
                                            auto.AutoPost(ship, invoice);
                                            break;
                                    }

                                    if (invoice.PO_NAV_Status == "PO_NOT_FOUND")
                                    {
                                        DateTime dateChecker;
                                        try { dateChecker = DateTime.Parse(invoice.InvoiceDate); }
                                        catch { dateChecker = invoice.ReceiveDate; }

                                        DateTime date = Constants.GetNumberBusinessOfDays(dateChecker, Constants.PoNotFoundDays);
                                        if (DateTime.Now.Date > date.Date)
                                            invoice.InvoiceRetry = false;
                                        else
                                            invoice.InvoiceRetry = true;
                                    }
                                    Database.UpdateInvoice(invoice.InvoiceID, invoice.PDFFileName, invoice.InvoiceRetry, invoice.Kwiktagged);
                                }
                                catch(Exception ex)
                                {
                                    SqlCommand cmd = new SqlCommand("Invoice Error");
                                    cmd.Parameters.Add(new SqlParameter("@invoiceID", invoice.InvoiceID));
                                    cmd.Parameters.Add(new SqlParameter("@dir", dir));
                                    Constants.ERRORS.Add(new Error(ex, cmd, "Main", "Program"));
                                }
                            }

                            new Save_Batch_Excel_Report(vendor.Invoices.ToArray(), ship.ShipType, Constants.BatchFolder + batch.BatchNumber, vendor.VendorName);
                        }
                    }

                    foreach (Batch batch in ship.PoNotFoundBatches)
                    {
                        foreach (Vendor vendor in batch.Vendors)
                        {
                            foreach (InvoiceHeader invoice in vendor.Invoices)
                            {
                                string dir = "";
                                try
                                {
                                    dir = Constants.InvoiceDropFolder + @"\" + ship.ShipType + @"\" + Constants.BatchFolder + @"PO_NOT_FOUND\" + Constants.BatchFolder + batch.BatchNumber + @"\" + vendor.VendorName + @"\";
                                    CheckDirectory(dir);
                                    if (invoice.PDFFileName.Length == 0 || !File.Exists(invoice.PDFFileName))
                                    {
                                        BuildSingleInvoicePDF b = new BuildSingleInvoicePDF(invoice, dir, Constants.InvoiceDropFolder + @"\" + Constants.MasterBatchFolder);
                                        invoice.PDFFileName = b.CreatePDF();
                                        if (invoice.PDFFileName.Length == 0)
                                            invoice.Errors.Add("PDF File for invoice is empty!");
                                        //Email.SendEmail("Please see attached", invoice.Vendor + " eInvoice " + invoice.InvoiceID, "", "invoicing@govsci.com", "", "", invoice.PDFFileName, true);
                                    }

                                    DateTime dateChecker;
                                    try { dateChecker = DateTime.Parse(invoice.InvoiceDate); }
                                    catch { dateChecker = invoice.ReceiveDate; }

                                    DateTime date = Constants.GetNumberBusinessOfDays(dateChecker, Constants.PoNotFoundDays);
                                    if (DateTime.Now.Date > date.Date || invoice.PO_NAV_Status == "PO_POSTED")
                                        invoice.InvoiceRetry = false;
                                    else
                                        invoice.InvoiceRetry = true;

                                    Database.UpdateInvoice(invoice.InvoiceID, invoice.PDFFileName, invoice.InvoiceRetry, invoice.Kwiktagged);
                                }
                                catch(Exception ex)
                                {
                                    SqlCommand cmd = new SqlCommand("Invoice Error");
                                    cmd.Parameters.Add(new SqlParameter("@invoiceID", invoice.InvoiceID));
                                    cmd.Parameters.Add(new SqlParameter("@dir", dir));
                                    Constants.ERRORS.Add(new Error(ex, cmd, "Main", "Program"));
                                }
                            }

                            new Save_Batch_Excel_Report(vendor.Invoices.ToArray(), ship.ShipType, Constants.BatchFolder + @"PO_NOT_FOUND\" + Constants.BatchFolder + batch.BatchNumber, vendor.VendorName);
                        }
                    }
                }
            }

            return ships;
        }        

        private void CheckDirectory(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        private string CheckDirectory(string orgpath, int i)
        {
            string path = orgpath;

            if (i != 0)
            {
                if (orgpath.EndsWith("\\"))
                    path = orgpath.Remove(orgpath.LastIndexOf("\\"));
                path = $"{path}_{i}\\";
            }

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                return path;
            }
            else
                path = CheckDirectory(orgpath, ++i);

            return path;
        }

        private List<Batch> CheckBatchList(List<Batch> batches, int batchNo, InvoiceHeader invoice)
        {
            bool addBatch = false, addVendor = false;
            Batch batch = batches.Find(b => b.BatchNumber == batchNo);
            if (batch == null)
            {
                batch = new Batch(batchNo);
                addBatch = true;
            }

            Vendor vendor = batch.Vendors.Find(v => v.VendorName == invoice.Vendor);
            if (vendor == null)
            {
                vendor = new Vendor(invoice.Vendor);
                addVendor = true;
            }

            InvoiceHeader inv = vendor.Invoices.Find(i => i.OrderID == invoice.OrderID);
            if (inv == null)
            {
                vendor.Invoices.Add(invoice);
                vendor.Total += invoice.InvoiceTotal;

                if (addVendor) batch.Vendors.Add(vendor);
                if (addBatch) batches.Add(batch);

                return batches;
            }
            else
            {
                batchNo += 1;
                return CheckBatchList(batches, batchNo, invoice);
            }
        }

        private int CheckBatchFolder(string folder, string orderID, int batch, string vendor)
        {
            string dir = folder + @"\" + Constants.BatchFolder + batch + @"\" + vendor + @"\";
            CheckDirectory(dir);
            bool exists = false;
            foreach (string file in Directory.GetFiles(dir))
                if (file.Contains(orderID)) exists = true;

            if (exists)
            {
                batch = batch + 1;
                return CheckBatchFolder(folder, orderID, batch, vendor);
            }
            else
                return batch;
        }
        private bool CheckFile(string folder, string invoiceNumber)
        {
            foreach (string file in Directory.GetFiles(folder))
                if (file.Contains(invoiceNumber)) return true;

            return false;
        }

        public static void Main(string[] args)
        {
            DateTime appStarted = DateTime.Now;
            Constants.ERRORS = new List<Error>();
            try
            {
                new Program(DateTime.Now.AddDays(-5), appStarted);
            }
            catch(Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, "Program", "Main"));
            }

            if (Constants.ERRORS.Count > 0)
            {
                if (!ReportSent && ReportNeedstoBeSent)
                {
                    SendReport report = new SendReport(invoices, ships, invoicesOnHold, emptyVendorInvoices, appStarted);
                    if (report.ReportErrors.Count > 0)
                    {
                        Constants.ERRORS.AddRange(report.ReportErrors);
                        new SendCsvReport(invoices, ships, invoicesOnHold, emptyVendorInvoices, appStarted);
                    }
                }
                Email.SendErrorMessage();
            }
        }
    }
}
