using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electronic_Invoice_Report.Objects
{
    public class InvoiceCount
    {
        public InvoiceCount(string vendorName, string format)
        {
            VendorName = vendorName;
            Format = format;
        }

        public string VendorName { get; }
        public string Format { get; }
        public int Count { get; set; }
    }
}
