using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
{
    public class VendorCounts
    {
        public VendorCounts(string vendorName, DateTime invoiceDate)
        {
            VendorName = vendorName;
            InvoiceDate = invoiceDate;
            Ships = new List<Ship>();
            InvoicesOnHold = new List<InvoiceHeader>();
        }

        public string VendorName { get; }
        public DateTime InvoiceDate { get; }

        public List<Ship> Ships { get; }
        public List<InvoiceHeader> InvoicesOnHold { get; }
    }
}
