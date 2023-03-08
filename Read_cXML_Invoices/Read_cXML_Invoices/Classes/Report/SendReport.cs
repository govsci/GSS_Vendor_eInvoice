using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using System.Xml;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Read_cXML_Invoices.Objects;


namespace Read_cXML_Invoices.Classes.Report
{
    public class SendReport
    {
        private HSSFWorkbook workbook;
        public List<Objects.Error> ReportErrors = new List<Objects.Error>();
        public List<Invoice> invoicesReceived;

        public SendReport(List<InvoiceHeader> invoices, List<Ship> ships, List<InvoiceHeader> invoicesOnHold, List<InvoiceHeader> emptyVendorInvoices, DateTime appStarted)
        {
            invoicesReceived = GetDailySupplierInvoices.GetThem(appStarted);

            try
            {
                InitializeWorkbook();
                
                if (invoices.Count > 0 || invoicesOnHold.Count > 0)
                {
                    AddBatchInfoWorksheet(ships); //Batch Information

                    AddCountsWorksheet(ships, invoicesOnHold); //EIR Counts
                    decimal batchTotal = AddInvoicesReceivedWorksheet(invoices); //Invoices Received Report

                    invoicesOnHold = invoicesOnHold.OrderBy(i => i.Vendor).ThenBy(i => i.ReleaseDate).ToList();
                    AddInvoicesOnHoldWorksheet(invoicesOnHold); //Invoices On Hold Report

                    AddHandsOnInvoicesWorksheet(ships);
                    AddPoNotFoundWorksheet(ships);

                    AddPostedInvoicesWorksheet(ships);
                    AddPostedInvoicesFailedWorksheet(ships);

                    if (emptyVendorInvoices.Count > 0)
                        AddEmptyVendorInvoicesWorksheet(emptyVendorInvoices);

                    AddDSIRCountsWorksheet(invoicesReceived); //DSIR Counts
                    AddDSIRWorksheet(invoicesReceived); //DSIR Details

                    string excelPath = Constants.ReportPath + DateTime.Now.ToString(@"yyyy\\MM\\dd\\");
                    if (!Directory.Exists(excelPath)) Directory.CreateDirectory(excelPath);

                    excelPath += DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".InvoiceReport.xls";

                    WriteToFile(excelPath);
                    workbook.Close();

                    SendTheReport(excelPath, batchTotal, ships);
                }
                else
                {
                    workbook.Close();
                    throw new Exception("Invoices are empty");
                }
            }
            catch (Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "SendReport"));
            }
        }

        private void InitializeWorkbook()
        {
            workbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "Government Scientific Source, Inc.";
            workbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "Invoices Received Report";
            workbook.SummaryInformation = si;
        }

        //WORKSHEETS
        private void AddBatchInfoWorksheet(List<Ship> ships)
        {
            try
            {
                ISheet worksheet = workbook.CreateSheet("Batch Information");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("Batch Information");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "Batch Information";

                IRow headerRow2 = worksheet.CreateRow(1);
                ICell headerCell2 = headerRow2.CreateCell(0);
                headerCell2.SetCellValue("Normal Batches");

                //header
                IRow hRow = worksheet.CreateRow(2);
                hRow.CreateCell(0).SetCellValue("Ship Type");
                hRow.CreateCell(1).SetCellValue("Batch Number");
                hRow.CreateCell(2).SetCellValue("Vendor");
                hRow.CreateCell(3).SetCellValue("dABatchID");
                hRow.CreateCell(4).SetCellValue("# of Invoices");
                hRow.CreateCell(5).SetCellValue("Total Amount");

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

                int rowNo = 3;
                foreach (Ship ship in ships)
                {
                    foreach (Batch batch in ship.Batches)
                    {
                        foreach (Vendor vendor in batch.Vendors)
                        {
                            try
                            {
                                IRow row = worksheet.CreateRow(rowNo);
                                row.CreateCell(0).SetCellValue(ship.ShipType);
                                row.CreateCell(1).SetCellValue(batch.BatchNumber.ToString());
                                row.CreateCell(2).SetCellValue(vendor.VendorName);
                                row.CreateCell(3).SetCellValue(vendor.daBatchId);
                                row.CreateCell(4).SetCellValue(vendor.Invoices.Count.ToString());
                                row.CreateCell(5).SetCellValue(vendor.Total.ToString("G29"));

                                rowNo++;
                            }
                            catch (Exception ex)
                            {
                                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddBatchInfoWorksheet"));
                            }
                        }
                    }
                }

                rowNo += 1;

                IRow headerRow3 = worksheet.CreateRow(rowNo);
                ICell headerCell3 = headerRow3.CreateCell(0);
                headerCell3.SetCellValue("PO Not Found Batches");

                rowNo += 1;

                IRow hRow2 = worksheet.CreateRow(rowNo);
                hRow2.CreateCell(0).SetCellValue("Ship Type");
                hRow2.CreateCell(1).SetCellValue("Batch Number");
                hRow2.CreateCell(2).SetCellValue("Vendor");
                hRow2.CreateCell(3).SetCellValue("# of Invoices");
                hRow2.CreateCell(4).SetCellValue("Total Amount");

                foreach (ICell cell in hRow2.Cells)
                    cell.CellStyle = style1;

                rowNo += 1;

                foreach (Ship ship in ships)
                {
                    foreach (Batch batch in ship.PoNotFoundBatches)
                    {
                        foreach (Vendor vendor in batch.Vendors)
                        {
                            try
                            {
                                IRow row = worksheet.CreateRow(rowNo);
                                row.CreateCell(0).SetCellValue(ship.ShipType);
                                row.CreateCell(1).SetCellValue(batch.BatchNumber.ToString());
                                row.CreateCell(2).SetCellValue(vendor.VendorName);
                                row.CreateCell(3).SetCellValue(vendor.Invoices.Count.ToString());
                                row.CreateCell(4).SetCellValue(vendor.Total.ToString("G29"));

                                rowNo++;
                            }
                            catch (Exception ex)
                            {
                                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddBatchInfoWorksheet"));
                            }
                        }
                    }
                }

                for (int i = 0; i < 19; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddBatchInfoWorksheet"));
            }
        }
        private void AddCountsWorksheet(List<Ship> ships, List<InvoiceHeader> invoicesOnHold)
        {
            try
            {
                invoicesOnHold = invoicesOnHold.OrderBy(i => i.Vendor).ThenBy(i => i.InvoiceDate).ToList();
                List<VendorCounts> vendors = CalculateCounts(ships, invoicesOnHold);

                ISheet worksheet = workbook.CreateSheet("EIR Counts");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("Electronic Invoices Received Counts");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "Electronic Invoices Received Counts";

                //header
                IRow hRow = worksheet.CreateRow(1);
                hRow.CreateCell(0).SetCellValue("Vendor Name");
                hRow.CreateCell(1).SetCellValue("Invoice Date");

                string columnString = "Vendor Name,Invoice Date";
                StringBuilder csv = new StringBuilder();

                List<CountColumn> columns = new List<CountColumn>();

                int column = 2;
                foreach (VendorCounts vendor in vendors)
                {
                    foreach (Ship ship in vendor.Ships)
                    {
                        try
                        {
                            foreach (Batch batch in ship.Batches)
                            {
                                string columnName = $"{ship.ShipType} - Batch {batch.BatchNumber}";
                                CountColumn cc = columns.Find(c => c.ShipType == ship.ShipType && c.BatchNumber == batch.BatchNumber && !c.PO_Not_Found);
                                if (cc == null)
                                {
                                    columns.Add(new CountColumn(ship.ShipType, batch.BatchNumber, column, false));
                                    hRow.CreateCell(column).SetCellValue(columnName);
                                    column = column + 1;

                                    columnString += "," + columnName;
                                }
                            }

                            foreach (Batch batch in ship.PoNotFoundBatches)
                            {
                                string columnName = $"{ship.ShipType} - PO Not Found - Batch {batch.BatchNumber}";
                                CountColumn cc = columns.Find(c => c.ShipType == ship.ShipType && c.BatchNumber == batch.BatchNumber && c.PO_Not_Found);
                                if (cc == null)
                                {
                                    columns.Add(new CountColumn(ship.ShipType, batch.BatchNumber, column, true));
                                    hRow.CreateCell(column).SetCellValue(columnName);
                                    column = column + 1;

                                    columnString += "," + columnName;
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddCountsWorksheet"));
                        }
                    }
                }

                columns.Add(new CountColumn("Invoices On Hold", 0, column, false));
                hRow.CreateCell(column).SetCellValue("Invoices On Hold");

                columnString += ",Invoices On Hold";

                column = column + 1;
                columns.Add(new CountColumn("Total", 0, column, false));
                hRow.CreateCell(column).SetCellValue("Total");

                columnString += ",Total";
                csv.Append(columnString);

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

                foreach (VendorCounts vendor in vendors)
                {
                    try
                    {
                        int count = 0;
                        IRow row = worksheet.CreateRow(rowNo);

                        row.CreateCell(0).SetCellValue(vendor.VendorName);
                        row.CreateCell(1).SetCellValue(vendor.InvoiceDate.ToShortDateString());


                        foreach (Ship ship in vendor.Ships)
                        {
                            foreach (Batch batch in ship.Batches)
                            {
                                CountColumn col = columns.Find(c => c.ShipType == ship.ShipType && c.BatchNumber == batch.BatchNumber && !c.PO_Not_Found);
                                row.CreateCell(col.ColumnID).SetCellValue(batch.Invoices.Count);
                                count += batch.Invoices.Count;
                            }

                            foreach (Batch batch in ship.PoNotFoundBatches)
                            {
                                CountColumn col = columns.Find(c => c.ShipType == ship.ShipType && c.BatchNumber == batch.BatchNumber && c.PO_Not_Found);
                                row.CreateCell(col.ColumnID).SetCellValue(batch.Invoices.Count);
                                count += batch.Invoices.Count;
                            }
                        }

                        CountColumn col1 = columns.Find(c => c.ShipType == "Invoices On Hold");
                        row.CreateCell(col1.ColumnID).SetCellValue(vendor.InvoicesOnHold.Count);
                        count += vendor.InvoicesOnHold.Count;

                        col1 = columns.Find(c => c.ShipType == "Total");
                        row.CreateCell(col1.ColumnID).SetCellValue(count);

                        rowNo++;
                    }
                    catch (Exception ex)
                    {
                        ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddCountsWorksheet"));
                    }
                }

                for (int i = 0; i < column + 2; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddCountsWorksheet"));
            }
        }
        private decimal AddInvoicesReceivedWorksheet(List<InvoiceHeader> invoices)
        {
            decimal batchTotal = 0.0M;

            try
            {
                ISheet worksheet = workbook.CreateSheet("Invoices Received Report");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("Invoices Received Report");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "Invoices Received Report Invoices";

                //header
                IRow hRow = worksheet.CreateRow(1);
                hRow.CreateCell(0).SetCellValue("Vendor Name");
                hRow.CreateCell(1).SetCellValue("Invoice ID");
                hRow.CreateCell(2).SetCellValue("Invoice Date");
                hRow.CreateCell(3).SetCellValue("PO #");
                hRow.CreateCell(4).SetCellValue("PO Date");
                hRow.CreateCell(5).SetCellValue("Subtotal Amount");
                hRow.CreateCell(6).SetCellValue("Tax");
                hRow.CreateCell(7).SetCellValue("Shipping Amount");
                hRow.CreateCell(8).SetCellValue("Special Handling Amount");
                hRow.CreateCell(9).SetCellValue("Invoice Total");
                hRow.CreateCell(10).SetCellValue("Line Number");
                hRow.CreateCell(11).SetCellValue("Part Number");
                hRow.CreateCell(12).SetCellValue("Description");
                hRow.CreateCell(13).SetCellValue("Unit of Measure");
                hRow.CreateCell(14).SetCellValue("Quantity");
                hRow.CreateCell(15).SetCellValue("Unit Price");
                hRow.CreateCell(16).SetCellValue("Line Tax");
                hRow.CreateCell(17).SetCellValue("Line Total");
                hRow.CreateCell(18).SetCellValue("Receive Date");

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
                foreach (InvoiceHeader invoice in invoices)
                {
                    foreach (InvoiceLine line in invoice.Lines)
                    {
                        try
                        {
                            IRow row = worksheet.CreateRow(rowNo);

                            string invoiceTotal = "";
                            if (invoice.DueAmount > 0.0M)
                                invoiceTotal = invoice.DueAmount.ToString("G29");
                            else if (invoice.NetAmount > 0.00M)
                                invoiceTotal = invoice.NetAmount.ToString("G29");
                            else if (invoice.GrossAmount > 0.00M)
                                invoiceTotal = invoice.GrossAmount.ToString("G29");

                            row.CreateCell(0).SetCellValue(invoice.Vendor);
                            row.CreateCell(1).SetCellValue(invoice.InvoiceID);
                            row.CreateCell(2).SetCellValue(invoice.InvoiceDate);
                            row.CreateCell(3).SetCellValue(invoice.OrderID);
                            row.CreateCell(4).SetCellValue(invoice.OrderDate);
                            row.CreateCell(5).SetCellValue(invoice.SubTotalAmount.ToString("G29"));
                            row.CreateCell(6).SetCellValue(invoice.Tax.ToString("G29"));
                            row.CreateCell(7).SetCellValue(invoice.ShippingAmount.ToString("G29"));
                            row.CreateCell(8).SetCellValue(invoice.SpecialHandlingAmount.ToString("G29"));
                            row.CreateCell(9).SetCellValue(invoiceTotal);
                            row.CreateCell(10).SetCellValue(line.LineNumber);
                            row.CreateCell(11).SetCellValue(line.SupplierPartID);
                            row.CreateCell(12).SetCellValue(line.Description);
                            row.CreateCell(13).SetCellValue(line.UnitOfMeasure);
                            row.CreateCell(14).SetCellValue(line.Quantity.ToString("G29"));
                            row.CreateCell(15).SetCellValue(line.UnitPrice.ToString("G29"));
                            row.CreateCell(16).SetCellValue(line.Tax.ToString("G29"));
                            row.CreateCell(17).SetCellValue(line.LineTotal.ToString("G29"));
                            row.CreateCell(18).SetCellValue(invoice.ReceiveDate.ToString("MM/dd/yyyy hh:mm tt"));

                            batchTotal += (line.Quantity * line.UnitPrice);

                            rowNo++;
                        }
                        catch (Exception ex)
                        {
                            ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddInvoicesReceivedWorksheet"));
                        }
                    }
                }

                for (int i = 0; i < 19; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddInvoicesReceivedWorksheet"));
            }

            return batchTotal;
        }
        private void AddHandsOnInvoicesWorksheet(List<Ship> ships)
        {
            try
            {
                int count = 0;

                foreach (Ship ship in ships)
                    foreach (Batch batch in ship.PoNotFoundBatches)
                        foreach (Vendor vendor in batch.Vendors)
                            foreach (InvoiceHeader invoice in vendor.Invoices)
                                if (!invoice.InvoiceRetry && invoice.PO_NAV_Status != "PO_POSTED")
                                    count++;

                if (count > 0)
                {
                    ISheet worksheet = workbook.CreateSheet("Hands On Required Invoices Report");

                    IRow headerRow = worksheet.CreateRow(0);
                    ICell headerCell = headerRow.CreateCell(1);
                    headerCell.SetCellValue("Hands On Required Invoices Report");

                    worksheet.Header.Left = HSSFHeader.Page;
                    worksheet.Header.Center = "Hands On Required Invoices";

                    //header
                    IRow hRow = worksheet.CreateRow(1);
                    hRow.CreateCell(0).SetCellValue("Ship Type");
                    hRow.CreateCell(1).SetCellValue("Batch No.");
                    hRow.CreateCell(2).SetCellValue("Vendor Name");
                    hRow.CreateCell(3).SetCellValue("Invoice ID");
                    hRow.CreateCell(4).SetCellValue("Invoice Date");
                    hRow.CreateCell(5).SetCellValue("Received Date");
                    hRow.CreateCell(6).SetCellValue("PO #");
                    hRow.CreateCell(7).SetCellValue("PO Date");
                    hRow.CreateCell(8).SetCellValue("Subtotal Amount");
                    hRow.CreateCell(9).SetCellValue("Tax");
                    hRow.CreateCell(10).SetCellValue("Shipping");
                    hRow.CreateCell(11).SetCellValue("Special Handling");
                    hRow.CreateCell(12).SetCellValue("Invoice Total");
                    hRow.CreateCell(13).SetCellValue("Status Code");

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
                    foreach (Ship ship in ships)
                    {
                        foreach (Batch batch in ship.PoNotFoundBatches)
                        {
                            foreach (Vendor vendor in batch.Vendors)
                            {
                                foreach (InvoiceHeader invoice in vendor.Invoices)
                                {
                                    try
                                    {
                                        if (!invoice.InvoiceRetry && invoice.PO_NAV_Status != "PO_POSTED")
                                        {
                                            IRow row = worksheet.CreateRow(rowNo);

                                            string invoiceTotal = "";
                                            if (invoice.DueAmount > 0.0M)
                                                invoiceTotal = invoice.DueAmount.ToString("G29");
                                            else if (invoice.NetAmount > 0.00M)
                                                invoiceTotal = invoice.NetAmount.ToString("G29");
                                            else if (invoice.GrossAmount > 0.00M)
                                                invoiceTotal = invoice.GrossAmount.ToString("G29");

                                            row.CreateCell(0).SetCellValue(ship.ShipType);
                                            row.CreateCell(1).SetCellValue(batch.BatchNumber);
                                            row.CreateCell(2).SetCellValue(vendor.VendorName);
                                            row.CreateCell(3).SetCellValue(invoice.InvoiceID);
                                            row.CreateCell(4).SetCellValue(invoice.InvoiceDate);
                                            row.CreateCell(5).SetCellValue(invoice.ReceiveDate.ToShortDateString());
                                            row.CreateCell(6).SetCellValue(invoice.OrderID);
                                            row.CreateCell(7).SetCellValue(invoice.OrderDate);
                                            row.CreateCell(8).SetCellValue(invoice.SubTotalAmount.ToString("G29"));
                                            row.CreateCell(9).SetCellValue(invoice.Tax.ToString("G29"));
                                            row.CreateCell(10).SetCellValue(invoice.ShippingAmount.ToString("G29"));
                                            row.CreateCell(11).SetCellValue(invoice.SpecialHandlingAmount.ToString("G29"));
                                            row.CreateCell(12).SetCellValue(invoice.InvoiceTotal.ToString("G29"));
                                            row.CreateCell(13).SetCellValue(invoice.PO_NAV_Status);

                                            rowNo++;
                                        }
                                    }
                                    catch(Exception ex)
                                    {
                                        ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddHandsOnInvoicesWorksheet"));
                                    }
                                }
                            }
                        }
                    }

                    for (int i = 0; i < 14; i++)
                    {
                        worksheet.AutoSizeColumn(i);
                        GC.Collect();
                    }
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddHandsOnInvoicesWorksheet"));
            }
        }
        private void AddPoNotFoundWorksheet(List<Ship> ships)
        {
            try
            {
                ISheet worksheet = workbook.CreateSheet("PO Not Found Invoices Report");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("PO Not Found Invoices Report");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "PO Not Found Invoices Invoices";

                //header
                IRow hRow = worksheet.CreateRow(1);
                hRow.CreateCell(0).SetCellValue("Ship Type");
                hRow.CreateCell(1).SetCellValue("Batch No.");
                hRow.CreateCell(2).SetCellValue("Vendor Name");
                hRow.CreateCell(3).SetCellValue("Invoice ID");
                hRow.CreateCell(4).SetCellValue("Invoice Date");
                hRow.CreateCell(5).SetCellValue("Received Date");
                hRow.CreateCell(6).SetCellValue("PO #");
                hRow.CreateCell(7).SetCellValue("PO Date");
                hRow.CreateCell(8).SetCellValue("Subtotal Amount");
                hRow.CreateCell(9).SetCellValue("Tax");
                hRow.CreateCell(10).SetCellValue("Shipping");
                hRow.CreateCell(11).SetCellValue("Special Handling");
                hRow.CreateCell(12).SetCellValue("Invoice Total");
                hRow.CreateCell(13).SetCellValue("Retry?");
                hRow.CreateCell(14).SetCellValue("Status Code");

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
                foreach (Ship ship in ships)
                {
                    foreach (Batch batch in ship.PoNotFoundBatches)
                    {
                        foreach (Vendor vendor in batch.Vendors)
                        {
                            foreach (InvoiceHeader invoice in vendor.Invoices)
                            {
                                try
                                {
                                    IRow row = worksheet.CreateRow(rowNo);

                                    string invoiceTotal = "";
                                    if (invoice.DueAmount > 0.0M)
                                        invoiceTotal = invoice.DueAmount.ToString("G29");
                                    else if (invoice.NetAmount > 0.00M)
                                        invoiceTotal = invoice.NetAmount.ToString("G29");
                                    else if (invoice.GrossAmount > 0.00M)
                                        invoiceTotal = invoice.GrossAmount.ToString("G29");

                                    row.CreateCell(0).SetCellValue(ship.ShipType);
                                    row.CreateCell(1).SetCellValue(batch.BatchNumber);
                                    row.CreateCell(2).SetCellValue(vendor.VendorName);
                                    row.CreateCell(3).SetCellValue(invoice.InvoiceID);
                                    row.CreateCell(4).SetCellValue(invoice.InvoiceDate);
                                    row.CreateCell(5).SetCellValue(invoice.ReceiveDate.ToShortDateString());
                                    row.CreateCell(6).SetCellValue(invoice.OrderID);
                                    row.CreateCell(7).SetCellValue(invoice.OrderDate);
                                    row.CreateCell(8).SetCellValue(invoice.SubTotalAmount.ToString("G29"));
                                    row.CreateCell(9).SetCellValue(invoice.Tax.ToString("G29"));
                                    row.CreateCell(10).SetCellValue(invoice.ShippingAmount.ToString("G29"));
                                    row.CreateCell(11).SetCellValue(invoice.SpecialHandlingAmount.ToString("G29"));
                                    row.CreateCell(12).SetCellValue(invoice.InvoiceTotal.ToString("G29"));
                                    row.CreateCell(13).SetCellValue(invoice.InvoiceRetry ? "Yes" : "No");
                                    row.CreateCell(14).SetCellValue(invoice.PO_NAV_Status);

                                    rowNo++;
                                }
                                catch(Exception ex)
                                {
                                    ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddPoNotFoundWorksheet"));
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < 15; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddPoNotFoundWorksheet"));
            }
        }
        private void AddInvoicesOnHoldWorksheet(List<InvoiceHeader> invoices)
        {
            try
            {
                ISheet worksheet = workbook.CreateSheet("NDS Invoices On Hold Report");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("NDS Invoices On Hold Report");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "NDS Invoices On Hold Report Invoices";

                //header
                IRow hRow = worksheet.CreateRow(1);
                hRow.CreateCell(0).SetCellValue("Vendor Name");
                hRow.CreateCell(1).SetCellValue("Invoice ID");
                hRow.CreateCell(2).SetCellValue("Invoice Date");
                hRow.CreateCell(3).SetCellValue("PO #");
                hRow.CreateCell(4).SetCellValue("PO Date");
                hRow.CreateCell(5).SetCellValue("Invoice Total");
                hRow.CreateCell(6).SetCellValue("Date to Release to DocAlpha");

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
                foreach (InvoiceHeader invoice in invoices)
                {
                    try
                    {
                        IRow row = worksheet.CreateRow(rowNo);

                        string invoiceTotal = "";
                        if (invoice.DueAmount > 0.0M)
                            invoiceTotal = invoice.DueAmount.ToString("G29");
                        else if (invoice.NetAmount > 0.00M)
                            invoiceTotal = invoice.NetAmount.ToString("G29");
                        else if (invoice.GrossAmount > 0.00M)
                            invoiceTotal = invoice.GrossAmount.ToString("G29");

                        DateTime dateChecker;
                        try { dateChecker = DateTime.Parse(invoice.InvoiceDate); }
                        catch { dateChecker = invoice.ReceiveDate; }

                        row.CreateCell(0).SetCellValue(invoice.Vendor);
                        row.CreateCell(1).SetCellValue(invoice.InvoiceID);
                        row.CreateCell(2).SetCellValue(dateChecker.ToShortDateString());
                        row.CreateCell(3).SetCellValue(invoice.OrderID);
                        row.CreateCell(4).SetCellValue(invoice.OrderDate);
                        row.CreateCell(5).SetCellValue(invoiceTotal);
                        row.CreateCell(6).SetCellValue(invoice.ReleaseDate.ToShortDateString());

                        rowNo++;
                    }
                    catch(Exception ex)
                    {
                        ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddInvoicesOnHoldWorksheet"));
                    }
                }

                for (int i = 0; i < 7; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddInvoicesOnHoldWorksheet"));
            }
        }
        private void AddPostedInvoicesWorksheet(List<Ship> ships)
        {
            try
            {
                ISheet worksheet = workbook.CreateSheet("Posted Purchase Orders Report");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("Posted Purchase Orders Report");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "Posted Purchase Orders";

                //header
                IRow hRow = worksheet.CreateRow(1);
                hRow.CreateCell(0).SetCellValue("Ship Type");
                hRow.CreateCell(1).SetCellValue("Batch No.");
                hRow.CreateCell(2).SetCellValue("Vendor Name");
                hRow.CreateCell(3).SetCellValue("dA Batch ID");
                hRow.CreateCell(4).SetCellValue("Invoice ID");
                hRow.CreateCell(5).SetCellValue("Invoice Date");
                hRow.CreateCell(6).SetCellValue("Received Date");
                hRow.CreateCell(7).SetCellValue("PO #");
                hRow.CreateCell(8).SetCellValue("PO Date");
                hRow.CreateCell(9).SetCellValue("Subtotal Amount");
                hRow.CreateCell(10).SetCellValue("Tax");
                hRow.CreateCell(11).SetCellValue("GL Accounts Amount");
                hRow.CreateCell(12).SetCellValue("Invoice Total");
                hRow.CreateCell(13).SetCellValue("Purchase Order Posted Total");
                hRow.CreateCell(14).SetCellValue("Totals Match?");
                hRow.CreateCell(15).SetCellValue("PO Posted Receipt");
                hRow.CreateCell(16).SetCellValue("PO Posted Invoice");
                hRow.CreateCell(17).SetCellValue("Kwiktagged?");
                hRow.CreateCell(18).SetCellValue("Notes");

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
                                    try
                                    {
                                        IRow row = worksheet.CreateRow(rowNo);

                                        decimal specialAmountTotal = 0.0M;
                                        foreach (var shipLine in invoice.ShipLines)
                                            specialAmountTotal += (shipLine.Quantity * shipLine.UnitPrice);

                                        row.CreateCell(0).SetCellValue(ship.ShipType);
                                        row.CreateCell(1).SetCellValue(batch.BatchNumber.ToString());
                                        row.CreateCell(2).SetCellValue(vendor.VendorName);
                                        row.CreateCell(3).SetCellValue(vendor.daBatchId);
                                        row.CreateCell(4).SetCellValue(invoice.InvoiceID);
                                        row.CreateCell(5).SetCellValue(invoice.InvoiceDate);
                                        row.CreateCell(6).SetCellValue(invoice.ReceiveDate.ToShortDateString());
                                        row.CreateCell(7).SetCellValue(invoice.OrderID);
                                        row.CreateCell(8).SetCellValue(invoice.OrderDate);
                                        row.CreateCell(9).SetCellValue(invoice.SubTotalAmount.ToString("G29"));
                                        row.CreateCell(10).SetCellValue(invoice.Tax.ToString("G29"));
                                        row.CreateCell(11).SetCellValue(specialAmountTotal.ToString("G29"));
                                        row.CreateCell(12).SetCellValue(invoice.CalculatedInvoiceTotal.ToString("G29"));
                                        row.CreateCell(13).SetCellValue(invoice.PurchaseOrderPostedTotal.ToString("G29"));
                                        row.CreateCell(14).SetCellValue((invoice.CalculatedInvoiceTotal == invoice.PurchaseOrderPostedTotal) ? "Yes" : "No");
                                        row.CreateCell(15).SetCellValue(invoice.PurchaseOrderPostedReceipt ? "Yes" : "No");
                                        row.CreateCell(16).SetCellValue(invoice.PurchaseOrderPostedInvoice ? "Yes" : "No");
                                        row.CreateCell(17).SetCellValue(invoice.Kwiktagged ? "Yes" : "No");
                                        row.CreateCell(18).SetCellValue(invoice.Notes);

                                        rowNo++;

                                    }
                                    catch (Exception ex)
                                    {
                                        ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddPostedInvoicesWorksheet"));
                                    }
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < 19; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddPostedInvoicesWorksheet"));
            }
        }
        private void AddPostedInvoicesFailedWorksheet(List<Ship> ships)
        {
            try
            {
                ISheet worksheet = workbook.CreateSheet("Invoices With Errors Report");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("Invoices With Errors Report");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "Invoices With Errors";

                //header
                IRow hRow = worksheet.CreateRow(1);
                hRow.CreateCell(0).SetCellValue("Ship Type");
                hRow.CreateCell(1).SetCellValue("Batch No.");
                hRow.CreateCell(2).SetCellValue("Vendor Name");
                hRow.CreateCell(3).SetCellValue("Invoice ID");
                hRow.CreateCell(4).SetCellValue("Invoice Date");
                hRow.CreateCell(5).SetCellValue("Received Date");
                hRow.CreateCell(6).SetCellValue("PO #");
                hRow.CreateCell(7).SetCellValue("PO Date");
                hRow.CreateCell(8).SetCellValue("Subtotal Amount");
                hRow.CreateCell(9).SetCellValue("Tax");
                hRow.CreateCell(10).SetCellValue("GL Accounts Amount");
                hRow.CreateCell(11).SetCellValue("Invoice Total");
                hRow.CreateCell(12).SetCellValue("Purchase Order Posted Total");
                hRow.CreateCell(13).SetCellValue("Totals Match?");
                hRow.CreateCell(14).SetCellValue("PO Posted Receipt");
                hRow.CreateCell(15).SetCellValue("PO Posted Invoice");
                hRow.CreateCell(16).SetCellValue("Kwiktagged");
                hRow.CreateCell(17).SetCellValue("Errors");

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

                foreach (var ship in ships)
                {
                    foreach (var batch in ship.Batches)
                    {
                        foreach (var vendor in batch.Vendors)
                        {
                            foreach (var invoice in vendor.Invoices)
                            {
                                //if (invoice.Errors.Count > 0 || invoice.ErrorCode.Length > 0)
                                if (invoice.Errors.Count > 0)
                                {
                                    try
                                    {
                                        IRow row = worksheet.CreateRow(rowNo);

                                        decimal specialAmountTotal = 0.0M;
                                        foreach (var shipLine in invoice.ShipLines)
                                            specialAmountTotal += (shipLine.Quantity * shipLine.UnitPrice);

                                        string errors = "";
                                        foreach (string error in invoice.Errors)
                                            errors += error + Environment.NewLine;
                                        //if (invoice.ErrorCode.Length > 0)
                                        //    errors = errors.Length > 0 ? invoice.ErrorCode + Environment.NewLine + errors : invoice.ErrorCode;

                                        row.CreateCell(0).SetCellValue(ship.ShipType);
                                        row.CreateCell(1).SetCellValue(batch.BatchNumber.ToString());
                                        row.CreateCell(2).SetCellValue(invoice.Vendor);
                                        row.CreateCell(3).SetCellValue(invoice.InvoiceID);
                                        row.CreateCell(4).SetCellValue(invoice.InvoiceDate);
                                        row.CreateCell(5).SetCellValue(invoice.ReceiveDate.ToShortDateString());
                                        row.CreateCell(6).SetCellValue(invoice.OrderID);
                                        row.CreateCell(7).SetCellValue(invoice.OrderDate);
                                        row.CreateCell(8).SetCellValue(invoice.SubTotalAmount.ToString("G29"));
                                        row.CreateCell(9).SetCellValue(invoice.Tax.ToString("G29"));
                                        row.CreateCell(10).SetCellValue(specialAmountTotal.ToString("G29"));
                                        row.CreateCell(11).SetCellValue(invoice.CalculatedInvoiceTotal.ToString("G29"));
                                        row.CreateCell(12).SetCellValue(invoice.PurchaseOrderPostedTotal.ToString("G29"));
                                        row.CreateCell(13).SetCellValue((invoice.CalculatedInvoiceTotal == invoice.PurchaseOrderPostedTotal) ? "Yes" : "No");
                                        row.CreateCell(14).SetCellValue(invoice.PurchaseOrderPostedReceipt ? "Yes" : "No");
                                        row.CreateCell(15).SetCellValue(invoice.PurchaseOrderPostedInvoice ? "Yes" : "No");
                                        row.CreateCell(16).SetCellValue(invoice.Kwiktagged ? "Yes" : "No");
                                        try { row.CreateCell(17).SetCellValue(errors); }
                                        catch { row.CreateCell(17).SetCellValue("Errors were too long for this field"); Console.WriteLine("Invoice ID {0} Errors: {1}", invoice.InvoiceID, errors); }

                                        row.Cells[15].CellStyle.WrapText = true;

                                        rowNo++;
                                    }
                                    catch (Exception ex)
                                    {
                                        ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddPostedInvoicesFailedWorksheet"));
                                    }
                                }
                            }
                        }
                    }
                }

                for (int i = 0; i < 18; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch (Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddPostedInvoicesFailedWorksheet"));
            }
        }
        private void AddEmptyVendorInvoicesWorksheet(List<InvoiceHeader> invoices)
        {
            try
            {
                ISheet worksheet = workbook.CreateSheet("Empty Vendor Invoices Report");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("Empty Vendor Invoices Report");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "Empty Vendor Invoices Report";

                //header
                IRow hRow = worksheet.CreateRow(1);
                hRow.CreateCell(0).SetCellValue("Vendor Name");
                hRow.CreateCell(1).SetCellValue("Invoice ID");
                hRow.CreateCell(2).SetCellValue("Invoice Date");
                hRow.CreateCell(3).SetCellValue("PO #");
                hRow.CreateCell(4).SetCellValue("PO Date");
                hRow.CreateCell(5).SetCellValue("Subtotal Amount");
                hRow.CreateCell(6).SetCellValue("Tax");
                hRow.CreateCell(7).SetCellValue("Shipping Amount");
                hRow.CreateCell(8).SetCellValue("Special Handling Amount");
                hRow.CreateCell(9).SetCellValue("Invoice Total");
                hRow.CreateCell(10).SetCellValue("Receive Date");

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
                foreach (InvoiceHeader invoice in invoices)
                {
                    try
                    {
                        IRow row = worksheet.CreateRow(rowNo);

                        string invoiceTotal = "";
                        if (invoice.DueAmount > 0.0M)
                            invoiceTotal = invoice.DueAmount.ToString("G29");
                        else if (invoice.NetAmount > 0.00M)
                            invoiceTotal = invoice.NetAmount.ToString("G29");
                        else if (invoice.GrossAmount > 0.00M)
                            invoiceTotal = invoice.GrossAmount.ToString("G29");

                        row.CreateCell(0).SetCellValue(invoice.Vendor);
                        row.CreateCell(1).SetCellValue(invoice.InvoiceID);
                        row.CreateCell(2).SetCellValue(invoice.InvoiceDate);
                        row.CreateCell(3).SetCellValue(invoice.OrderID);
                        row.CreateCell(4).SetCellValue(invoice.OrderDate);
                        row.CreateCell(5).SetCellValue(invoice.SubTotalAmount.ToString("G29"));
                        row.CreateCell(6).SetCellValue(invoice.Tax.ToString("G29"));
                        row.CreateCell(7).SetCellValue(invoice.ShippingAmount.ToString("G29"));
                        row.CreateCell(8).SetCellValue(invoice.SpecialHandlingAmount.ToString("G29"));
                        row.CreateCell(9).SetCellValue(invoiceTotal);
                        row.CreateCell(10).SetCellValue(invoice.ReceiveDate.ToString("MM/dd/yyyy hh:mm tt"));

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
                        ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddEmptyVendorInvoicesWorksheet"));
                    }
                }

                for (int i = 0; i < 11; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddEmptyVendorInvoicesWorksheet"));
            }
        }
        private void AddDSIRCountsWorksheet(List<Invoice> invoices)
        {
            try
            {
                List<InvoiceCount> summaries = CalculateInvoiceCounts(invoices);

                ISheet worksheet = workbook.CreateSheet("DSIR Counts");

                IRow headerRow = worksheet.CreateRow(0);
                ICell headerCell = headerRow.CreateCell(1);
                headerCell.SetCellValue("Daily Supplier Invoices Received Counts");

                worksheet.Header.Left = HSSFHeader.Page;
                worksheet.Header.Center = "Daily Supplier Invoices Received Counts";

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
                        ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddDSIRCountsWorksheet"));
                    }
                }

                for (int i = 0; i < 3; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddDSIRCountsWorksheet"));
            }
        }
        private void AddDSIRWorksheet(List<Invoice> invoices)
        {
            try
            {
                ISheet worksheet = workbook.CreateSheet("DSIR Details");

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
                hRow.CreateCell(9).SetCellValue("Uploaded To Table");
                hRow.CreateCell(10).SetCellValue("DocAlpha Date");

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
                            row.CreateCell(8).SetCellValue(invoice.Format == "XML" || invoice.Format == "EDI" ? invoice.InvoiceReceived.ToString("MM/dd/yyyy hh:mm tt") : "");
                            row.CreateCell(9).SetCellValue(invoice.InTable);
                            row.CreateCell(10).SetCellValue(invoice.DocAlphaDate);

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
                            ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddDSIRWorksheet"));
                        }
                    }
                }

                for (int i = 0; i < 11; i++)
                {
                    worksheet.AutoSizeColumn(i);
                    GC.Collect();
                }
            }
            catch(Exception ex)
            {
                ReportErrors.Add(new Objects.Error(ex, "SendReport", "AddDSIRWorksheet"));
            }
        }

        //Misc Functions
        private List<VendorCounts> CalculateCounts(List<Ship> ships, List<InvoiceHeader> invoicesOnHold)
        {
            List<VendorCounts> vendors = new List<VendorCounts>();

            foreach(Ship ship in ships)
            {
                foreach(Batch batch in ship.Batches)
                {
                    foreach(Vendor vendor in batch.Vendors)
                    {
                        foreach(InvoiceHeader invoice in vendor.Invoices)
                        {
                            DateTime invoiceDate = DateTime.Parse(invoice.InvoiceDate);
                            VendorCounts tempvendor = vendors.Find(v => v.VendorName == vendor.VendorName && v.InvoiceDate.Date == invoiceDate.Date);
                            if (tempvendor == null)
                            {
                                tempvendor = new VendorCounts(vendor.VendorName, invoiceDate);
                                vendors.Add(tempvendor);
                            }

                            Ship tempship = tempvendor.Ships.Find(s => s.ShipType == ship.ShipType);
                            if(tempship == null)
                            {
                                tempship = new Ship(ship.ShipType);
                                tempvendor.Ships.Add(tempship);
                            }

                            Batch tempbatch = tempship.Batches.Find(b => b.BatchNumber == batch.BatchNumber);
                            if(tempbatch == null)
                            {
                                tempbatch = new Batch(batch.BatchNumber);
                                tempship.Batches.Add(tempbatch);
                            }

                            InvoiceHeader tempinvoice = tempbatch.Invoices.Find(i => i.InvoiceID == invoice.InvoiceID);
                            if (tempinvoice == null)
                                tempbatch.Invoices.Add(invoice);
                        }
                    }
                }

                foreach (Batch batch in ship.PoNotFoundBatches)
                {
                    foreach (Vendor vendor in batch.Vendors)
                    {
                        foreach (InvoiceHeader invoice in vendor.Invoices)
                        {
                            DateTime invoiceDate = DateTime.Parse(invoice.InvoiceDate);
                            VendorCounts tempvendor = vendors.Find(v => v.VendorName == vendor.VendorName && v.InvoiceDate.Date == invoiceDate.Date);
                            if (tempvendor == null)
                            {
                                tempvendor = new VendorCounts(vendor.VendorName, invoiceDate);
                                vendors.Add(tempvendor);
                            }

                            Ship tempship = tempvendor.Ships.Find(s => s.ShipType == ship.ShipType);
                            if (tempship == null)
                            {
                                tempship = new Ship(ship.ShipType);
                                tempvendor.Ships.Add(tempship);
                            }

                            Batch tempbatch = tempship.PoNotFoundBatches.Find(b => b.BatchNumber == batch.BatchNumber);
                            if (tempbatch == null)
                            {
                                tempbatch = new Batch(batch.BatchNumber);
                                tempship.PoNotFoundBatches.Add(tempbatch);
                            }

                            InvoiceHeader tempinvoice = tempbatch.Invoices.Find(i => i.InvoiceID == invoice.InvoiceID);
                            if (tempinvoice == null)
                                tempbatch.Invoices.Add(invoice);
                        }
                    }
                }
            }

            foreach(InvoiceHeader invoice in invoicesOnHold)
            {
                DateTime invoicedate = DateTime.Parse(invoice.InvoiceDate);
                VendorCounts tempvendor = vendors.Find(v => v.VendorName == invoice.Vendor && v.InvoiceDate.Date == invoicedate.Date);
                if(tempvendor == null)
                {
                    tempvendor = new VendorCounts(invoice.Vendor, invoicedate);
                    vendors.Add(tempvendor);
                }

                tempvendor.InvoicesOnHold.Add(invoice);
            }

            vendors = vendors.OrderBy(v => v.VendorName).ThenByDescending(v => v.InvoiceDate.Date).ToList();

            return vendors;
        }
        private List<InvoiceCount> CalculateInvoiceCounts(List<Invoice> invoices)
        {
            List<InvoiceCount> counters = new List<InvoiceCount>();

            foreach (Invoice invoice in invoices)
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
        
        private void WriteToFile(string excelPath)
        {
            FileStream file = new FileStream(excelPath, FileMode.Create);
            workbook.Write(file);
            file.Close();
        }        

        private void SendTheReport(string excelPath, decimal batchTotal, List<Ship> ships)
        {
            string msg = "Please review the attached Invoice Report for invoices that were received today, " + DateTime.Now.ToString(@"MM/dd/yyyy") + ".<br><br>";

            msg += "Batch Information (Ship Type, Batch Number, Vendor, and Total) as been moved to the Excel file.";

            msg += "<br><br>Total is: $" + batchTotal.ToString("G29");
            Email.SendEmail(msg, "Electronic Invoices Received", "", Constants.EmailRecipients, "", "", excelPath, true);
        }

        private string Convert(string path)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var wb = app.Workbooks.Open(path);
            wb.SaveAs(Filename: path + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();

            return path + "x";
        }

        private class CountColumn
        {
            public CountColumn(string shipType, int batchNumber, int columnID, bool notFound)
            {
                ShipType = shipType;
                BatchNumber = batchNumber;
                ColumnID = columnID;
                PO_Not_Found = notFound;
            }

            public string ShipType { get; }
            public int BatchNumber { get; }
            public bool PO_Not_Found { get; }
            public int ColumnID { get; }
        }
    }
}
