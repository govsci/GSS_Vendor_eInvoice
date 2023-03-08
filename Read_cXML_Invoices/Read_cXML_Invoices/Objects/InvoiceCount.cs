using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
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
