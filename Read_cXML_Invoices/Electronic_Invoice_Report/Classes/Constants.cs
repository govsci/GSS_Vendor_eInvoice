using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utilities;
using System.Text.RegularExpressions;

namespace Electronic_Invoice_Report.Classes
{
    public static class Constants
    {
        public static string cXMLInvoiceFolder = @"C:\Sean\Dump\Others\Invoice\Incoming\";
        public static void SendEmailNewVendor(string userAgent, string invoiceId)
        {
            Email.SendEmail("Please check the following details for new Electronic Invoice Vendor:<br /><br />Invoice ID: " + invoiceId + "<br />User Agent:" + userAgent, "New Electronic Invoice Vendor", "", "dev_error@govsci.com", "", "", "", true);
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
                v = Regex.Replace(v, @"\s{2,}"," ");
                v = v.Trim();
            }
            return v;
        }
    }
}
