using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes
{
    public class UploadData
    {
        private InvoiceHeader invoice;
        private bool include = true;
        private string query = "";

        public bool IncludeInvoice { get { return include; } }
        public string Query { get { return query; } }

        public UploadData(InvoiceHeader i)
        {
            invoice = i;
        }

        private bool CheckInvoice()
        {
            bool rVal = false;
            using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
            {
                dbcon.Open();
                query = "SELECT [invoiceID] FROM [dbo].[Ecommerce$Electronic Document Header] WHERE [invoiceID]='" + invoice.InvoiceID + "' AND [orderID] = '" + invoice.OrderID + "'";
                using (SqlDataReader rs = new SqlCommand(query, dbcon).ExecuteReader())
                    if (rs.Read())
                        rVal = true;
            }
            return rVal;
        }

        private string RemoveSpecialCharacters(string val)
        {
            return Regex.Replace(val, @"\s{2,}", "");
        }

        public void Upload()
        {
            SqlCommand cmd = null;
            if (!CheckInvoice())
            {
                try
                {
                    string extrinsics = "";
                    foreach (Extrinsic extrinsic in invoice.Extrinsics)
                        extrinsics += extrinsic.Name + ":" + extrinsic.Value + "|";
                    using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                    {
                        dbcon.Open();
                        cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@method", "INSERT INVOICE HEADER"));
                        cmd.Parameters.Add(new SqlParameter("@timestamp", invoice.Timestamp));
                        cmd.Parameters.Add(new SqlParameter("@payloadID", invoice.PayloadID));
                        cmd.Parameters.Add(new SqlParameter("@fromDomain", invoice.FromDomain));
                        cmd.Parameters.Add(new SqlParameter("@fromIdentity", invoice.FromIdentity));
                        cmd.Parameters.Add(new SqlParameter("@toDomain", invoice.ToDomain));
                        cmd.Parameters.Add(new SqlParameter("@toIdentity", invoice.ToIdentity));
                        cmd.Parameters.Add(new SqlParameter("@senderDomain", invoice.SenderDomain));
                        cmd.Parameters.Add(new SqlParameter("@senderIdentity", invoice.SenderIdentity));
                        cmd.Parameters.Add(new SqlParameter("@sharedSecret", invoice.SharedSecret));
                        cmd.Parameters.Add(new SqlParameter("@userAgent", invoice.UserAgent));
                        cmd.Parameters.Add(new SqlParameter("@invoiceID", invoice.InvoiceID));
                        cmd.Parameters.Add(new SqlParameter("@purpose", invoice.Purpose));
                        cmd.Parameters.Add(new SqlParameter("@operation", invoice.Operation));
                        cmd.Parameters.Add(new SqlParameter("@invoiceDate", invoice.InvoiceDate));
                        cmd.Parameters.Add(new SqlParameter("@paymentTermDays", invoice.PaymentTermNumberOfDays));
                        cmd.Parameters.Add(new SqlParameter("@paymentTermPercent", invoice.PaymentTermPercentRate));
                        cmd.Parameters.Add(new SqlParameter("@extrinsics", extrinsics));
                        cmd.Parameters.Add(new SqlParameter("@orderID", invoice.OrderID));
                        cmd.Parameters.Add(new SqlParameter("@documentRefrPayloadID", invoice.DocumentReferencePayloadID));
                        cmd.Parameters.Add(new SqlParameter("@orderDate", invoice.OrderDate));
                        cmd.Parameters.Add(new SqlParameter("@subtotalAmount", invoice.SubTotalAmount));
                        cmd.Parameters.Add(new SqlParameter("@tax", invoice.Tax));
                        cmd.Parameters.Add(new SqlParameter("@grossAmount", invoice.GrossAmount));
                        cmd.Parameters.Add(new SqlParameter("@netAmount", invoice.NetAmount));
                        cmd.Parameters.Add(new SqlParameter("@dueAmount", invoice.DueAmount));
                        cmd.Parameters.Add(new SqlParameter("@shippingAmount", invoice.ShippingAmount));
                        cmd.Parameters.Add(new SqlParameter("@specialHandlingAmount", invoice.SpecialHandlingAmount));
                        cmd.Parameters.Add(new SqlParameter("@invoiceDetailDiscount", invoice.InvoiceDetailDiscount));
                        cmd.Parameters.Add(new SqlParameter("@invoiceTotal", invoice.InvoiceTotal));
                        cmd.Parameters.Add(new SqlParameter("@vendor", invoice.Vendor));

                        int id = 0;
                        using (SqlDataReader rs = cmd.ExecuteReader()) if (rs.Read()) id = rs.GetInt32(0);

                        if (id > 0)
                        {

                            foreach (AddressObject role in invoice.Roles)
                            {
                                cmd = new SqlCommand(cmd.CommandText, cmd.Connection);
                                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                                cmd.Parameters.Add(new SqlParameter("@method", "INSERT INVOICE ROLE"));
                                cmd.Parameters.Add(new SqlParameter("@parent", id));
                                cmd.Parameters.Add(new SqlParameter("@role", role.Role));
                                cmd.Parameters.Add(new SqlParameter("@addressID", role.ID));
                                cmd.Parameters.Add(new SqlParameter("@name", role.Name));
                                cmd.Parameters.Add(new SqlParameter("@deliverTo", role.DeliverTo));
                                cmd.Parameters.Add(new SqlParameter("@street", role.Street));
                                cmd.Parameters.Add(new SqlParameter("@city", role.City));
                                cmd.Parameters.Add(new SqlParameter("@state", role.State));
                                cmd.Parameters.Add(new SqlParameter("@postalCode", role.PostalCode));
                                cmd.Parameters.Add(new SqlParameter("@countryCode", role.CountryCode));
                                cmd.Parameters.Add(new SqlParameter("@country", role.Country));
                                cmd.ExecuteNonQuery();
                            }

                            foreach (InvoiceLine line in invoice.Lines)
                            {
                                cmd = new SqlCommand(cmd.CommandText, cmd.Connection);
                                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                                cmd.Parameters.Add(new SqlParameter("@method", "INSERT INVOICE LINE"));
                                cmd.Parameters.Add(new SqlParameter("@parent", id));
                                cmd.Parameters.Add(new SqlParameter("@invoiceID", invoice.InvoiceID));
                                cmd.Parameters.Add(new SqlParameter("@lineNumber", line.LineNumber));
                                cmd.Parameters.Add(new SqlParameter("@quantity", line.Quantity));
                                cmd.Parameters.Add(new SqlParameter("@unitOfMeasure", line.UnitOfMeasure));
                                cmd.Parameters.Add(new SqlParameter("@unitPrice", line.UnitPrice));
                                cmd.Parameters.Add(new SqlParameter("@refrLineNumber", line.ReferenceLineNumber));
                                cmd.Parameters.Add(new SqlParameter("@supplierPartId", line.SupplierPartID));
                                cmd.Parameters.Add(new SqlParameter("@description", line.Description));
                                cmd.Parameters.Add(new SqlParameter("@lineTotal", line.LineTotal));
                                cmd.Parameters.Add(new SqlParameter("@shipLine", line.ShipLine));
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Constants.ERRORS.Add(new Error(ex, cmd, "UploadData", "Upload"));
                }
            }
            else
                include = false;
        }
    }
}
