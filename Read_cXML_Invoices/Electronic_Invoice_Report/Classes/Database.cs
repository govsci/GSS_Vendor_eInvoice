using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Utilities;
using System.Configuration;
using Electronic_Invoice_Report.Objects;

namespace Electronic_Invoice_Report.Classes
{
    public static class Database
    {
        public static EmailConfig GetEmailConfiguration()
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(DatabaseConnectionStrings.PrdEcomDb))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.Get.Email.Configuration]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET GET EMAIL"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                            return new EmailConfig(rs["host"].ToString(), rs["username"].ToString(), rs["password"].ToString(), rs["domain"].ToString(), rs["admin"].ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                Email.SendErrorMessage(ex, "Electronic_Invoice_Report", "Database.GetEmailConfiguration()", cmd);
            }

            return new EmailConfig(ConfigurationManager.AppSettings["emailDownHost"], ConfigurationManager.AppSettings["emailUsername"], ConfigurationManager.AppSettings["emailPassword"], ConfigurationManager.AppSettings["emailDomain"], ConfigurationManager.AppSettings["emailAdmin"]);
        }

        public static List<EmailFolderConfig> GetEmailFolderConfigs()
        {
            List<EmailFolderConfig> folders = new List<EmailFolderConfig>();
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(DatabaseConnectionStrings.PrdEcomDb))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Applications.Download.Process.Email.Factory.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET EMAIL FOLDER CONFIG"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        while (rs.Read())
                        {
                            if (rs["documentType"].ToString().ToUpper() == "INVOICE")
                                folders.Add(new EmailFolderConfig(
                                    rs["documentType"].ToString(),
                                    rs["emailFolder"].ToString(),
                                    rs["localFolder"].ToString()
                                    ));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Email.SendErrorMessage(ex, "Electronic_Invoice_Report", "Database.GetEmailFolderConfigs()", cmd);
            }

            return folders;
        }

        public static void InsertInvoice(Invoice invoice)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(DatabaseConnectionStrings.PrdEcomDb))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "INSERT LOG"));
                    cmd.Parameters.Add(new SqlParameter("@documentType", "INVOICE"));
                    cmd.Parameters.Add(new SqlParameter("@format", invoice.Format));
                    cmd.Parameters.Add(new SqlParameter("@emailFrom", invoice.EmailFrom));
                    cmd.Parameters.Add(new SqlParameter("@emailSubject",invoice.EmailSubject));
                    cmd.Parameters.Add(new SqlParameter("@emailBody", invoice.EmailBody));
                    cmd.Parameters.Add(new SqlParameter("@fromDomain", invoice.FromDomain));
                    cmd.Parameters.Add(new SqlParameter("@fromIdentity", invoice.FromIdentity));
                    cmd.Parameters.Add(new SqlParameter("@toDomain", invoice.ToDomain));
                    cmd.Parameters.Add(new SqlParameter("@toIdentity", invoice.ToIdentity));
                    cmd.Parameters.Add(new SqlParameter("@senderDomain", invoice.SenderDomain));
                    cmd.Parameters.Add(new SqlParameter("@senderIdentity", invoice.SenderIdentity));
                    cmd.Parameters.Add(new SqlParameter("@sharedSecret", invoice.SharedSecret));
                    cmd.Parameters.Add(new SqlParameter("@userAgent", invoice.UserAgent));
                    cmd.Parameters.Add(new SqlParameter("@invoiceID", invoice.InvoiceID));
                    cmd.Parameters.Add(new SqlParameter("@invoiceReceived", invoice.InvoiceReceived));
                    cmd.Parameters.Add(new SqlParameter("@file", invoice.File));
                    cmd.Parameters.Add(new SqlParameter("@vendor", invoice.Vendor));
                    cmd.Parameters.Add(new SqlParameter("@orderID", invoice.OrderID));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                            invoice.PreviouslyLogged = rs.GetInt32(0);
                    }
                }
            }
            catch(Exception ex)
            {
                Email.SendErrorMessage(ex, "Electronic_Invoice_Report", "Database.InsertInvoice(Invoice invoice)", cmd);
            }
        }

        public static string GetVendor(string useragent, string fromid, string secret, string invoiceid, string remitToName)
        {
            string vendor = "";
            SqlCommand cmd = null;

            try
            {
                using (SqlConnection dbcon = new SqlConnection(DatabaseConnectionStrings.PrdEcomDb))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET VENDOR"));
                    cmd.Parameters.Add(new SqlParameter("@userAgent", useragent));
                    cmd.Parameters.Add(new SqlParameter("@name", remitToName));
                    cmd.Parameters.Add(new SqlParameter("@fromIdentity", fromid));
                    cmd.Parameters.Add(new SqlParameter("@sharedSecret", secret));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                            vendor = rs["vendor"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Email.SendErrorMessage(ex, "Electronic_Invoice_Report", "Database.GetVendor", null);
            }

            return vendor;
        }
    }
}
