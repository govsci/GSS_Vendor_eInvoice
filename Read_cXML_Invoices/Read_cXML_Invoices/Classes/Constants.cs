using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Text.RegularExpressions;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes
{
    public static class Constants
    {
        public static string DbConnectionEcommerce = ConfigurationManager.ConnectionStrings["EcommerceDb"].ConnectionString;
        public static string DbConnectionNavision = ConfigurationManager.ConnectionStrings["NavisionDb"].ConnectionString;

        public static string AppProfile = ConfigurationManager.AppSettings["AppProfile"];
        public static string InvoiceFolder = ConfigurationManager.AppSettings["InvoiceFolder"];
        public static string EdiInvoiceFolder = ConfigurationManager.AppSettings["EdiInvoiceFolder"];
        public static string InvoiceDropFolder = ConfigurationManager.AppSettings["InvoiceDropFolder"] + DateTime.Now.ToString(@"yyyy\\MM\\dd\\");
        public static string BatchFolder = ConfigurationManager.AppSettings["BatchFolder"];
        public static string MasterBatchFolder = ConfigurationManager.AppSettings["MasterBatchFolder"];
        public static string daBatchIDPreq = ConfigurationManager.AppSettings["daBatchIDPreq"];
        public static string ReportPath = ConfigurationManager.AppSettings["ReportPath"];
        public static string KwiktagURL = ConfigurationManager.AppSettings["KwiktagURL"];
        public static string KwiktagUserName = ConfigurationManager.AppSettings["KwiktagUserName"];
        public static string KwiktagPassword = ConfigurationManager.AppSettings["KwiktagPassword"];
        public static int NdsInvoiceDays = GetInt(ConfigurationManager.AppSettings["NdsInvoiceDays"], 7);
        public static int PoNotFoundDays = GetInt(ConfigurationManager.AppSettings["PoNotFoundDays"], 10);
        public static decimal InvoiceThresholdAmt = GetDecimal(ConfigurationManager.AppSettings["InvoiceThresholdAmt"], 5.00M);
        public static string EmailRecipients = ConfigurationManager.AppSettings["EmailRecipients"];
        public static List<string> RetryStatuses = new List<string> { "DOCUMENT_ID_MISSING", "PO_NOT_FOUND", "" };

        public static List<DayOfWeek> businessDays = new List<DayOfWeek> { DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday };
        public static List<Error> ERRORS;
        public static List<string> NewVendors;

        public static string CheckDropShip(AddressObject shipTo)
        {
            string dropShip = "";            

            if (shipTo != null)
            {
                string shipToName = shipTo.Name.ToUpper();
                string shipToStreet = shipTo.Street.ToUpper();
                string shipToDeliverTo = shipTo.DeliverTo.ToUpper();

                if (shipToName.Contains("GSS") || shipToName.Contains("GOVT SCIENTIFIC") || shipToName.Contains("GOVERNMENT SCIENTIFIC") || shipToName.Contains("GOV SCI") || shipToName.Contains("GOVERNMENT SCI") || shipToName.Contains("GOVMT SCIENTIFIC"))
                {
                    if (shipToStreet.Contains("BILL OF LADING") || shipToDeliverTo.Contains("BILL OF LADING"))
                        dropShip = "NONDROP_SHIP";
                    else if (shipTo.City == "Albuquerque" && shipTo.State == "NM" && shipTo.PostalCode.StartsWith("87107"))
                        dropShip = "NONDROP_SHIP";
                    else if (shipTo.City == "C" && shipTo.State == "NM" && shipTo.PostalCode.StartsWith("87107"))
                        dropShip = "NONDROP_SHIP";
                    else
                        dropShip = "DROP_SHIP";
                }

                if (shipToStreet.Contains("12351") && shipToStreet.Contains("SUNRISE") && shipToStreet.Contains("VALLEY") && (shipToStreet.Contains("DRIVE") || shipToStreet.Contains("DR")))
                    dropShip = "NONDROP_SHIP";
                else if (shipToStreet.Contains("258") && shipToStreet.Contains("LINDBERGH") && (shipToStreet.Contains("AVENUE") || shipToStreet.Contains("AVE") || shipToStreet.Contains("LANE") || shipToStreet.Contains("LN")))
                    dropShip = "NONDROP_SHIP";
                else if (shipToStreet.Contains("10903") && shipToStreet.Contains("MCBRIDE") && (shipToStreet.Contains("LANE") || shipToStreet.Contains("LN")))
                    dropShip = "NONDROP_SHIP";
                else if ((shipToStreet.Contains("2701") || shipToStreet.Contains("2701-C")) && shipToStreet.Contains("BROADWAY") && (shipToStreet.Contains("NE") || shipToStreet.Contains("STE C") || shipToStreet.Contains("SUITE C")))
                    dropShip = "NONDROP_SHIP";
                else if (shipToStreet.Contains("6724") && shipToStreet.Contains("PRESTON") && shipToStreet.Contains("SUITE C") && (shipToStreet.Contains("AVE") || shipToStreet.Contains("AVENUE")))
                    dropShip = "NONDROP_SHIP";
                else if (shipToStreet.Contains("13894") && shipToStreet.Contains("REDSKIN") && (shipToStreet.Contains("DRIVE") || shipToStreet.Contains("DR")))
                    dropShip = "NONDROP_SHIP";
                else
                    dropShip = "DROP_SHIP";
            }
            else
                dropShip = "DROP_SHIP";

            return dropShip;
        }

        public static string RemoveHtmlElements(string value)
        {
            string v = "";
            if (value != null)
            {
                v = Regex.Replace(value, @"\<.*?\>", string.Empty);
                v = Regex.Replace(v, @"\r|\n|\r\n", " ");
                v = Regex.Replace(v, @"\t", " ");
                v = v.Replace("_", "");
                v = v.Replace("This email has been scanned by the GSS Email Security Service.", "");
                v = v.Replace("If you have any questions, please email Support@govsci.com", "");
                v = v.Replace("&nbsp;", " ");
                v = Regex.Replace(v, @"\s{2,}", " ");
                v = v.Trim();
            }
            return v;
        }

        public static bool SkipInvoice(ref InvoiceHeader invoice)
        {
            if (invoice.ShipType == "NONDROP_SHIP")
            {
                DateTime dateChecker;
                try { dateChecker = DateTime.Parse(invoice.InvoiceDate); }
                catch { dateChecker = invoice.ReceiveDate; }

                DateTime date = GetNumberBusinessOfDays(dateChecker, NdsInvoiceDays);
                if (DateTime.Now.Date >= date.Date)
                    return true;
                else
                    invoice.ReleaseDate = date;
            }
            else
                return true;

            return false;
        }

        public static DateTime GetNumberBusinessOfDays(DateTime day, int numberDays)
        {
            int currentNumber = 0;
            for (int i = 0; i < numberDays; i++)
            {
                DateTime cur = day.AddDays(i);
                if (Constants.businessDays.Contains(cur.DayOfWeek))
                    currentNumber += 1;
                else if (cur.DayOfWeek == DayOfWeek.Saturday)
                    currentNumber += 3;
                else
                    currentNumber += 2;
            }

            DateTime release = day.AddDays(currentNumber);
            if (release.DayOfWeek == DayOfWeek.Saturday)
                release = release.AddDays(2);
            else if (release.DayOfWeek == DayOfWeek.Sunday)
                release = release.AddDays(1);

            return release;
        }

        public static int GetInt(string value, int defaultValue)
        {
            try { return int.Parse(value); }
            catch { return defaultValue; }
        }

        public static decimal GetDecimal(string value, decimal defaultValue)
        {
            try { return decimal.Parse(value); }
            catch { return defaultValue; }
        }
    }
}
