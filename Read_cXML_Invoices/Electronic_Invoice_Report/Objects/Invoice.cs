using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Electronic_Invoice_Report.Objects
{
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
        }

        public Invoice(string vendor, string format, string fromDomain, string fromId, string toDom, string toId, string senderDom, string senderId, string secret
            , string userAgent, string invoiceId, string orderId, DateTime invoiceReceived, string file)
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
    }
}
