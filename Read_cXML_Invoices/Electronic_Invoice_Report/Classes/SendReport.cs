using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Electronic_Invoice_Report.Objects;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using System.IO;
using Utilities;


namespace Electronic_Invoice_Report.Classes
{
    public static class SendReport
    {
        private static HSSFWorkbook workbook;

        public static void SendIt(List<Invoice> invoices)
        {
            invoices = invoices.OrderBy(i => i.Vendor).ThenBy(i => i.Format).ToList();

            InitializeWorkbook();
            AddCountsWorksheet(invoices);
            AddWorksheet(invoices);

            string excelpath = $@"C:\Sean\Dump\Others\Incoming\Reports\{DateTime.Now.ToString(@"yyyy\\MM\\dd\\")}";
            if (!Directory.Exists(excelpath)) Directory.CreateDirectory(excelpath);

            excelpath += DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".DailySupplierInvoicesReceived.xls";
            WriteToFile(excelpath);

            workbook.Close();
            SendEmail(excelpath);
        }

        private static void InitializeWorkbook()
        {
            workbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Government Scientific Source, Inc.";
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "Daily Supplier Invoices Received";
            workbook.SummaryInformation = si;
        }

        private static void AddWorksheet(List<Invoice> invoices)
        {
            ISheet worksheet = workbook.CreateSheet("Details");

            IRow headerRow = worksheet.CreateRow(0);
            ICell headerCell = headerRow.CreateCell(1);
            headerCell.SetCellValue("Daily Supplier Invoices Received");

            worksheet.Header.Left = HSSFHeader.Page;
            worksheet.Header.Center = "Daily Supplier Invoices Received Details";

            //header
            IRow hRow = worksheet.CreateRow(1);
            hRow.CreateCell(0).SetCellValue("Vendor Name");
            hRow.CreateCell(1).SetCellValue("Format");
            hRow.CreateCell(2).SetCellValue("From Email Address");
            hRow.CreateCell(3).SetCellValue("Email Subject");
            hRow.CreateCell(4).SetCellValue("Email Body");
            hRow.CreateCell(5).SetCellValue("Email Received Date");
            hRow.CreateCell(6).SetCellValue("Invoice ID");
            hRow.CreateCell(7).SetCellValue("Order ID");
            hRow.CreateCell(8).SetCellValue("XML Received Date");

            ICellStyle style1 = workbook.CreateCellStyle();
            var palette = workbook.GetCustomPalette();
            palette.SetColorAtIndex(57, 188, 214, 238);

            style1.FillForegroundColor = palette.GetColor(57).Indexed;
            style1.FillPattern = FillPattern.SolidForeground;

            style1.Alignment = HorizontalAlignment.Center;
            style1.WrapText = true;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            foreach (ICell cell in hRow.Cells)
                cell.CellStyle = style1;

            int rowNo = 2;
            foreach (Invoice invoice in invoices)
            {
                if (invoice.PreviouslyLogged == 0)
                {
                    try
                    {
                        IRow row = worksheet.CreateRow(rowNo);

                        string emailbody = Constants.RemoveHtmlElements(invoice.EmailBody);
                        row.CreateCell(0).SetCellValue(invoice.Vendor);
                        row.CreateCell(1).SetCellValue(invoice.Format);
                        row.CreateCell(2).SetCellValue(invoice.EmailFrom);
                        row.CreateCell(3).SetCellValue(invoice.EmailSubject);
                        row.CreateCell(4).SetCellValue(emailbody);
                        row.CreateCell(5).SetCellValue(invoice.Format == "EMAIL" ? invoice.InvoiceReceived.ToString("MM/dd/yyyy hh:mm tt") : "");
                        row.CreateCell(6).SetCellValue(invoice.InvoiceID);
                        row.CreateCell(7).SetCellValue(invoice.OrderID);
                        row.CreateCell(8).SetCellValue(invoice.Format == "XML" ? invoice.InvoiceReceived.ToString("MM/dd/yyyy hh:mm tt") : "");

                        ICellStyle style2 = workbook.CreateCellStyle();
                        style2.Alignment = HorizontalAlignment.Left;
                        style2.BorderBottom = BorderStyle.Thin;
                        style2.BorderLeft = BorderStyle.Thin;
                        style2.BorderRight = BorderStyle.Thin;
                        style2.BorderTop = BorderStyle.Thin;

                        foreach (ICell cell in row.Cells)
                            cell.CellStyle = style2;


                        rowNo++;
                    }
                    catch (Exception ex)
                    {
                        Email.SendErrorMessage(ex, "Electronic_Invoice_Report", "SendReport.AddWorksheet", null);
                    }
                }
            }

            for (int i = 0; i < 9; i++)
            {
                worksheet.AutoSizeColumn(i);
                GC.Collect();
            }
        }

        private static void AddCountsWorksheet(List<Invoice> invoices)
        {
            List<InvoiceCount> summaries = CalculateInvoiceCounts(invoices);

            ISheet worksheet = workbook.CreateSheet("Summary");

            IRow headerRow = worksheet.CreateRow(0);
            ICell headerCell = headerRow.CreateCell(1);
            headerCell.SetCellValue("Daily Supplier Invoices Received Summary");

            worksheet.Header.Left = HSSFHeader.Page;
            worksheet.Header.Center = "Daily Supplier Invoices Received Summary";

            //header
            IRow hRow = worksheet.CreateRow(1);
            hRow.CreateCell(0).SetCellValue("Vendor Name");
            hRow.CreateCell(1).SetCellValue("Format");
            hRow.CreateCell(2).SetCellValue("Count");

            ICellStyle style1 = workbook.CreateCellStyle();
            var palette = workbook.GetCustomPalette();
            palette.SetColorAtIndex(57, 188, 214, 238);

            style1.FillForegroundColor = palette.GetColor(57).Indexed;
            style1.FillPattern = FillPattern.SolidForeground;

            style1.Alignment = HorizontalAlignment.Center;
            style1.WrapText = true;
            style1.BorderBottom = BorderStyle.Thin;
            style1.BorderLeft = BorderStyle.Thin;
            style1.BorderRight = BorderStyle.Thin;
            style1.BorderTop = BorderStyle.Thin;

            foreach (ICell cell in hRow.Cells)
                cell.CellStyle = style1;

            int rowNo = 2;
            foreach (InvoiceCount summary in summaries)
            {
                try
                {
                    IRow row = worksheet.CreateRow(rowNo);

                    row.CreateCell(0).SetCellValue(summary.VendorName);
                    row.CreateCell(1).SetCellValue(summary.Format);
                    row.CreateCell(2).SetCellValue(summary.Count);

                    ICellStyle style2 = workbook.CreateCellStyle();
                    style2.Alignment = HorizontalAlignment.Left;
                    style2.BorderBottom = BorderStyle.Thin;
                    style2.BorderLeft = BorderStyle.Thin;
                    style2.BorderRight = BorderStyle.Thin;
                    style2.BorderTop = BorderStyle.Thin;

                    foreach (ICell cell in row.Cells)
                        cell.CellStyle = style2;

                    rowNo++;
                }
                catch (Exception ex)
                {
                    Email.SendErrorMessage(ex, "Electronic_Invoice_Report", "SendReport.AddCountsWorksheet", null);
                }
            }

            for (int i = 0; i < 3; i++)
            {
                worksheet.AutoSizeColumn(i);
                GC.Collect();
            }
        }

        private static List<InvoiceCount> CalculateInvoiceCounts(List<Invoice> invoices)
        {
            List<InvoiceCount> counters = new List<InvoiceCount>();

            foreach(Invoice invoice in invoices)
            {
                if (invoice.PreviouslyLogged == 0)
                {
                    InvoiceCount count = counters.Find(c => c.VendorName == invoice.Vendor && c.Format == invoice.Format);
                    if (count == null)
                    {
                        count = new InvoiceCount(invoice.Vendor, invoice.Format);
                        counters.Add(count);
                    }

                    count.Count = count.Count + 1;
                }
            }

            return counters;
        }

        private static void WriteToFile(string excelPath)
        {
            FileStream file = new FileStream(excelPath, FileMode.Create);
            workbook.Write(file);
            file.Close();
        }

        private static void SendEmail(string excelPath)
        {
            string msg = "Please review the attached Daily Supplier Invoices Received Report for invoices that were received recently, via Email and cXML.<br><br>"
                + "If any of these invoices are missing from today's DocAlpha batch, please inform IT at GSS-IT-Development@govsci.com.<br><br>"
                + "Please note, for invoices received via EMAIL, since they have not been parsed, only Vendor Name, From Email Address, Email Subject, Email Body, and Email Received Date will be displayed.<br>"
                + "For invoices received via XML, only Vendor Name, Invoice ID, Order ID, and XML Received Date will be displayed.";
            
            //Email.SendEmail(msg, "Daily Supplier Invoices Received", "", "ap@govsci.com", "gss-it-development@govsci.com", "", excelPath, true);
            Email.SendEmail(msg, "Daily Supplier Invoices Received", "", "zlingelbach@govsci.com", "", "", excelPath, true);
        }
    }
}
