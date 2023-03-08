using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
{
    public class Ship
    {
        public Ship(string ship)
        {
            ShipType = ship;
            Batches = new List<Batch>();
            PoNotFoundBatches = new List<Batch>();
        }

        public string ShipType { get; }
        public List<Batch> Batches { get; set; }
        public List<Batch> PoNotFoundBatches { get; set; }
    }

    public class Batch
    {
        public Batch(int batchNumber)
        {
            BatchNumber = batchNumber;
            Vendors = new List<Vendor>();
            Invoices = new List<InvoiceHeader>();
        }

        public int BatchNumber { get; }
        public List<Vendor> Vendors { get; set; }
        public List<InvoiceHeader> Invoices { get; }
    }

    public class Vendor
    {
        public Vendor(string vendor)
        {
            VendorName = vendor;
            Total = 0.0M;
            Invoices = new List<InvoiceHeader>();
            daBatchId = "";
        }

        public string VendorName { get; }
        public decimal Total { get; set; }
        public List<InvoiceHeader> Invoices { get; set; }
        public string daBatchId { get; set; }
    }
}
