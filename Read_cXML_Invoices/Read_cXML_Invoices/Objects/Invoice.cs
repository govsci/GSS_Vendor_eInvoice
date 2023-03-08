using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
{
    public class InvoiceHeader
    {
        public InvoiceHeader()
        {
            Roles = new List<AddressObject>();
            Extrinsics = new List<Extrinsic>();
            Lines = new List<InvoiceLine>();
            ShipLines = new List<InvoiceLine>();
            Errors = new List<string>();
            PurchaseOrderNo = "";
            CalculatedInvoiceTotal = 0.0M;
            PurchaseOrder_LineNos = new List<int>();
            PurchaseOrderPostedInvoice = false;
            PurchaseOrderPostedReceipt = false;
            PurchaseOrderPostedTotal = 0.0M;
            BuyFromVendorNo = "";
            PO_NAV_Status = "";
            ShipType = "";
            dABatchID = "";
            PO_Found = false;
            Notes = "";
            InvoiceRetry = false;
            Document_ID = "";
            //ErrorCode = "";
        }

        public string Timestamp { get; set; }
        public string PayloadID { get; set; }

        public string FromDomain { get; set; }
        public string FromIdentity { get; set; }
        public string ToDomain { get; set; }
        public string ToIdentity { get; set; }
        public string SenderDomain { get; set; }
        public string SenderIdentity { get; set; }
        public string SharedSecret { get; set; }
        public string UserAgent { get; set; }

        public string InvoiceID { get; set; }
        public string Purpose { get; set; }
        public string Operation { get; set; }
        public string InvoiceDate { get; set; }

        public List<AddressObject> Roles { get; set; }

        public string PaymentTermNumberOfDays { get; set; }
        public string PaymentTermPercentRate { get; set; }

        public List<Extrinsic> Extrinsics { get; }

        public string OrderID { get; set; }
        public string DocumentReferencePayloadID { get; set; }
        public string OrderDate { get; set; }

        public List<InvoiceLine> Lines { get; set; }
        public List<InvoiceLine> ShipLines { get; set; }

        public decimal SubTotalAmount { get; set; }
        public decimal Tax { get; set; }
        public decimal GrossAmount { get; set; }
        public decimal NetAmount { get; set; }
        public decimal DueAmount { get; set; }
        public decimal ShippingAmount { get; set; }
        public decimal SpecialHandlingAmount { get; set; }
        public decimal InvoiceTotal { get; set; }
        public decimal InvoiceDetailDiscount { get; set; }
        public string TrackingNo { get; set; }

        public DateTime ReceiveDate { get; set; }
        public string PDFFileName { get; set; }
        public string Vendor { get; set; }
        public bool PO_Found { get; set; }
        public string PO_NAV_Status { get; set; }
        public DateTime ReleaseDate { get; set; }
        public List<string> Errors { get; }
        //public string ErrorCode { get; set; }
        public string PurchaseOrderNo { get; set; }
        public string BuyFromVendorNo { get; set; }
        public string BuyFromVendorName { get; set; }
        public decimal CalculatedInvoiceTotal { get; set; }
        public List<int> PurchaseOrder_LineNos { get; set; }
        public bool PurchaseOrderPostedReceipt { get; set; }
        public bool PurchaseOrderPostedInvoice { get;set; }
        public decimal PurchaseOrderPostedTotal { get; set; }
        public string ShipType { get; set; }
        public string dABatchID { get; set; }
        public bool Kwiktagged { get; set; }
        public string Notes { get; set; }
        public bool InvoiceRetry { get; set; }
        public string Document_ID { get; set; }
    }
    public class InvoiceLine
    {
        public InvoiceLine() 
        {
            ShipLine = 0;
            GSS_Part_Number = "";
            PurchLine_LineNumber = 0;
        }

        public int LineNumber { get; set; }
        public decimal Quantity { get; set; }
        public string UnitOfMeasure { get; set; }
        public decimal UnitPrice { get; set; }
        public int ReferenceLineNumber { get; set; }
        public string SupplierPartID { get; set; }
        public string Description { get; set; }
        public decimal LineTotal { get; set; }
        public decimal Tax { get; set; }
        public int ShipLine { get; set; }
        public string GSS_Part_Number { get; set; }
        public int PurchLine_LineNumber { get; set; }
    }
    public class AddressObject
    {
        public AddressObject(string role, string id, string name, string deliverTo, string street, string city, string state, string postalCode, string countryCode, string country)
        {
            Role = role;
            ID = id;
            Name = name;
            DeliverTo = deliverTo;
            Street = street;
            City = city;
            State = state;
            PostalCode = postalCode;
            CountryCode = countryCode;
            Country = country;
        }
        public string Role { get; }
        public string ID { get; }
        public string Name { get; }
        public string DeliverTo { get; }
        public string Street { get; }
        public string City { get; }
        public string State { get; }
        public string PostalCode { get; }
        public string CountryCode { get; }
        public string Country { get; }
    }
    public class Extrinsic
    {
        public Extrinsic(string name, string value)
        {
            Name = name;
            Value = value;
        }
        public string Name { get; }
        public string Value { get; }
    }

    public class Invoice
    {
        public Invoice(string vendor, string format, string emailFrom, string emailSubject, string emailBody, DateTime emailDate)
        {
            Vendor = vendor;
            Format = format;
            EmailFrom = emailFrom;
            EmailSubject = emailSubject;
            EmailBody = emailBody;
            InvoiceReceived = emailDate;

            FromDomain = "";
            FromIdentity = "";
            ToDomain = "";
            ToIdentity = "";
            SenderDomain = "";
            SenderIdentity = "";
            SharedSecret = "";
            UserAgent = "";
            InvoiceID = "";
            OrderID = "";
            File = "";
            InTable = "";
            DocAlphaDate = "";
        }

        public Invoice(string vendor, string format, string fromDomain, string fromId, string toDom, string toId, string senderDom, string senderId, string secret
            , string userAgent, string invoiceId, string orderId, DateTime invoiceReceived, string file, string inTable, string docAlphaDate)
        {
            Vendor = vendor;
            Format = format;
            FromDomain = fromDomain;
            FromIdentity = fromId;
            ToDomain = toDom;
            ToIdentity = toId;
            SenderDomain = senderDom;
            SenderIdentity = senderId;
            SharedSecret = secret;
            UserAgent = userAgent;
            InvoiceID = invoiceId;
            OrderID = orderId;
            InvoiceReceived = invoiceReceived;
            File = file;
            InTable = inTable;
            DocAlphaDate = docAlphaDate;

            EmailFrom = "";
            EmailSubject = "";
            EmailBody = "";
        }

        public string Vendor { get; }
        public string Format { get; }
        public string EmailFrom { get; }
        public string EmailSubject { get; }
        public string EmailBody { get; }
        public string FromDomain { get; }
        public string FromIdentity { get; }
        public string ToDomain { get; }
        public string ToIdentity { get; }
        public string SenderDomain { get; }
        public string SenderIdentity { get; }
        public string SharedSecret { get; }
        public string UserAgent { get; }
        public string InvoiceID { get; }
        public string OrderID { get; }
        public DateTime InvoiceReceived { get; }
        public string File { get; }

        public int PreviouslyLogged { get; set; }
        public string InTable { get; }
        public string DocAlphaDate { get; }
        
    }
}
