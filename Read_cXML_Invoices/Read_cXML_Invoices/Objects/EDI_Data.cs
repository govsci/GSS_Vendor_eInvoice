using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
{
    public class EDI_Data
    {
        public EDI_Data(string po, string invDate, string ordDate, string amount)
        {
            PO = po;
            InvoiceDate = invDate;
            OrderDate = ordDate;
            Amount = amount;
        }

        public string PO { get; }
        public string InvoiceDate { get; }
        public string OrderDate { get; }
        public string Amount { get; }
    }
}
