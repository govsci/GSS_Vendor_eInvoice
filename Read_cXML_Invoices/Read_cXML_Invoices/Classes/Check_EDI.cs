using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes
{
    public class Check_EDI
    {
        private List<Invoice> invoices = new List<Invoice>();
        private DateTime AppStarted;

        public List<Invoice> Check(DateTime check, DateTime appStarted)
        {
            AppStarted = appStarted;
            invoices = new List<Invoice>();
            for (DateTime day = check.Date; day <= DateTime.Now.Date; day = day.AddDays(1))
                GetFiles(day);
            return invoices;
        }

        public void GetFiles(DateTime day)
        {
            string folder = Constants.EdiInvoiceFolder + $@"{day.ToString("yyyy")}\{day.ToString("MM")}\{day.ToString("dd")}\";
            if (Directory.Exists(folder))
            {
                string[] files = Directory.GetFiles(folder);
                foreach (string file in files)
                {
                    if (File.GetCreationTime(file) < AppStarted)
                    {
                        string[] fileNameArray = file.Split('\\');
                        string fileName = fileNameArray[fileNameArray.Length - 1];
                        if (fileName.StartsWith("810"))
                            ParseFile(file);
                    }
                }
            }
        }

        public void ParseFile(string file)
        {
            string contents = File.ReadAllText(file);
            string fromdomain = "", fromidentity = "", todomain = "", toidentity = "", invoiceId = "", orderId = "";

            Match match = Regex.Match(contents, @"ISA\*.{2}\*\s{10}\*.{2}\*\s{10}\*(?<fromdom>.{2})\*(?<fromid>.{15})\*(?<todom>.{2})\*(?<toid>.{15})\*");
            if (match.Success)
            {
                fromdomain = match.Groups["fromdom"].Value.Trim();
                fromidentity = match.Groups["fromid"].Value.Trim();
                todomain = match.Groups["todom"].Value.Trim();
                toidentity = match.Groups["toid"].Value.Trim();
            }

            string vendor = Database.GetVendor(
                fromidentity,
                fromidentity,
                "",
                "",
                "");

            MatchCollection matches = Regex.Matches(contents, @"~BIG\*\d{8}\*(?<invoiceid>.+?)\*\d{0,8}\*(?<pono>.+?)\*");
            foreach (Match m in matches)
            {
                invoiceId = m.Groups["invoiceid"].Value.Trim();
                orderId = m.Groups["pono"].Value.Trim();

                string checkTable = Database.CheckInvoice(invoiceId, orderId);

                invoices.Add(new Invoice(
                    vendor,
                    "EDI",
                    fromdomain,
                    fromidentity,
                    todomain,
                    toidentity,
                    "",
                    "",
                    "",
                    "",
                    invoiceId,
                    orderId,
                    File.GetCreationTime(file),
                    file,
                    checkTable == "NOT_FOUND" ? "NO" : "Yes",
                    checkTable != "0" && checkTable != "NOT_FOUND" ? checkTable : ""
                    ));
            }
        }
    }
}
