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
    public class Save_Batch_Excel_Report
    {
        private HSSFWorkbook workbook;

        public Save_Batch_Excel_Report(InvoiceHeader[] invoices, string ship, string batch, string vendor)
        {
            try
            {
                InitializeWorkbook();

                if (invoices.Length > 0)
                {
                    AddWorksheet(invoices, ship, batch, vendor);

                    string excelPath = Constants.InvoiceDropFolder + ship + "\\" + batch + "\\" + vendor + "\\";
                    if (!Directory.Exists(excelPath)) Directory.CreateDirectory(excelPath);

                    excelPath += DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".BatchReport.xls";

                    WriteToFile(excelPath);
                    workbook.Close();
                }
                else
                    workbook.Close();
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Objects.Error(ex, "Save_Batch_Excel_Report", "Save_Batch_Excel_Report"));
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

        private void AddWorksheet(InvoiceHeader[] invoices, string ship, string batch, string vendor)
        {
            decimal batchTotal = 0.0M;

            ISheet worksheet = workbook.CreateSheet("Batch Report");

            IRow headerRow = worksheet.CreateRow(0);
            ICell headerCell = headerRow.CreateCell(1);
            headerCell.SetCellValue(DateTime.Now.ToShortDateString() + " " + ship + " > " + batch + " > " + vendor + " Invoices Report");

            NPOI.SS.UserModel.IFont hFont = workbook.CreateFont();
            hFont.IsBold = true;
            hFont.FontHeightInPoints = 20;
            ICellStyle headerStyle = headerCell.CellStyle;
            headerStyle.SetFont(hFont);
            headerCell.CellStyle = headerStyle;

            worksheet.Header.Left = HSSFHeader.Page;
            worksheet.Header.Center = "Invoices Received Report" + " Invoices";

            //header
            IRow hRow = worksheet.CreateRow(1);
            hRow.CreateCell(0).SetCellValue("Invoice ID");
            hRow.CreateCell(1).SetCellValue("Invoice Date");
            hRow.CreateCell(2).SetCellValue("PO #");
            hRow.CreateCell(3).SetCellValue("PO Date");
            hRow.CreateCell(4).SetCellValue("# of Items");
            hRow.CreateCell(5).SetCellValue("Invoice Total");
            hRow.CreateCell(6).SetCellValue("Ship To Address");
            hRow.CreateCell(7).SetCellValue("Receive Date");

            ICellStyle style1 = workbook.CreateCellStyle();
            var palette = workbook.GetCustomPalette();
            palette.SetColorAtIndex(57, 188, 214, 238);

            style1.FillForegroundColor = palette.GetColor(57).Indexed;
            style1.FillPattern = FillPattern.SolidForeground;

            NPOI.SS.UserModel.IFont font1 = workbook.CreateFont();
            font1.IsBold = true;
            style1.SetFont(font1);

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
                    string shipToStr = "";
                    AddressObject shipTo = invoice.Roles.Find(s => s.Role == "shipTo");
                    if (shipTo != null)
                        shipToStr = shipTo.Name + " " + shipTo.DeliverTo + " " + shipTo.Street + " " + shipTo.City + ", " + shipTo.State + " " + shipTo.PostalCode + " " + (shipTo.CountryCode.Length > 0 ? shipTo.CountryCode : shipTo.Country);

                    IRow row = worksheet.CreateRow(rowNo);

                    row.CreateCell(0).SetCellValue(invoice.InvoiceID);
                    row.CreateCell(1).SetCellValue(invoice.InvoiceDate);
                    row.CreateCell(2).SetCellValue(invoice.OrderID);
                    row.CreateCell(3).SetCellValue(invoice.OrderDate);
                    row.CreateCell(4).SetCellValue(invoice.Lines.Count.ToString());
                    row.CreateCell(5).SetCellValue(invoice.InvoiceTotal.ToString("G29"));
                    row.CreateCell(6).SetCellValue(shipToStr);
                    row.CreateCell(7).SetCellValue(invoice.ReceiveDate.ToString("MM/dd/yyyy hh:mm tt"));

                    NPOI.SS.UserModel.IFont font2 = workbook.CreateFont();
                    font2.FontName = "Arial";
                    font2.IsBold = false;
                    font2.FontHeightInPoints = 10;

                    ICellStyle style2 = workbook.CreateCellStyle();
                    style2.Alignment = HorizontalAlignment.Left;
                    style2.BorderBottom = BorderStyle.Thin;
                    style2.BorderLeft = BorderStyle.Thin;
                    style2.BorderRight = BorderStyle.Thin;
                    style2.BorderTop = BorderStyle.Thin;
                    style2.SetFont(font2);

                    foreach (ICell cell in row.Cells)
                        cell.CellStyle = style2;

                    rowNo++;
                    batchTotal += invoice.InvoiceTotal;
                }
                catch (Exception ex)
                {
                    Constants.ERRORS.Add(new Objects.Error(ex, "Save_Batch_Excel_Report", "AddWorksheet"));
                }
            }

            IRow fRow = worksheet.CreateRow(rowNo);
            fRow.CreateCell(0).SetCellValue("");
            fRow.CreateCell(1).SetCellValue("");
            fRow.CreateCell(2).SetCellValue("");
            fRow.CreateCell(3).SetCellValue("");
            fRow.CreateCell(4).SetCellValue("Batch Total");
            fRow.CreateCell(5).SetCellValue(batchTotal.ToString("G29"));
            fRow.CreateCell(6).SetCellValue("");
            fRow.CreateCell(7).SetCellValue("");

            foreach (ICell cell in fRow.Cells)
                cell.CellStyle = style1;

            for (int i = 0; i < 8; i++)
            {
                worksheet.AutoSizeColumn(i);
                GC.Collect();
            }
        }

        private void WriteToFile(string excelPath)
        {
            FileStream file = new FileStream(excelPath, FileMode.Create);
            workbook.Write(file);
            file.Close();
        }

        private string Convert(string path)
        {
            var app = new Microsoft.Office.Interop.Excel.Application();
            var wb = app.Workbooks.Open(path);
            wb.SaveAs(Filename: path + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            wb.Close();
            app.Quit();

            //File.Delete(path);

            return path + "x";
        }
    }
}
