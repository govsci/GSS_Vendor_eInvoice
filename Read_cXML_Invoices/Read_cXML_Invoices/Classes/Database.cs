using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using Read_cXML_Invoices.Objects;

namespace Read_cXML_Invoices.Classes
{
    public static class Database
    {
        #region NAV_DB
        public static string GetItemNumber(string documentNo, string itemNo)
        {
            string no = "";
            SqlCommand cmd = null;
            try
            {
                string query = "[dbo].[Applications.NAV.Process.cXML.Documents.Control]";
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionNavision))
                {
                    dbcon.Open();
                    cmd = new SqlCommand(query, dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET ITEM NO"));
                    cmd.Parameters.Add(new SqlParameter("@documentID", documentNo));
                    cmd.Parameters.Add(new SqlParameter("@vendorItemNo", itemNo));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                            no = rs[0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "CheckDropShip"));
            }
            return no;
        }
        public static string InsertGlLineInvoice(string documentNo, string vendor, string no, string description, decimal unitCost, decimal quantity)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionNavision))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Applications.NAV.Process.cXML.Documents.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "INSERT GL LINE"));
                    cmd.Parameters.Add(new SqlParameter("@Document_No_Value", documentNo));
                    cmd.Parameters.Add(new SqlParameter("@Buyfrom_Vendor_No_Value", vendor));
                    cmd.Parameters.Add(new SqlParameter("@No_Value", no));
                    cmd.Parameters.Add(new SqlParameter("@DescriptionValue", description));
                    cmd.Parameters.Add(new SqlParameter("@Unit_Cost_LCYValue", unitCost));
                    cmd.Parameters.Add(new SqlParameter("@docAlpha_Direct_Unit_CostValue", unitCost));
                    cmd.Parameters.Add(new SqlParameter("@QuantityValue", quantity));
                    cmd.Parameters.Add(new SqlParameter("@docAlpha_Qty__to_ReceiveValue", quantity));
                    cmd.Parameters.Add(new SqlParameter("@docAlpha_Qty__to_InvoiceValue", quantity));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                        if (rs.Read())
                            return rs[0].ToString();
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "InsertGlLineInvoice"));
                return $"{ex.Message}";
            }

            return "";
        }
        public static decimal GetTotals(string documentNo, List<int> lineNos)
        {
            decimal total = 0.0M;
            SqlCommand cmd = null;
            StringBuilder lines = new StringBuilder();
            for (int i = 0; i < lineNos.Count; i++)
            {
                if (i == (lineNos.Count - 1))
                    lines.Append(lineNos[i]);
                else
                    lines.Append($"{lineNos[i]},");
            }

            try
            {
                string query = "[dbo].[Applications.NAV.Process.cXML.Documents.Control]";
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionNavision))
                {
                    dbcon.Open();
                    cmd = new SqlCommand(query, dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET TOTAL"));
                    cmd.Parameters.Add(new SqlParameter("@lineNos", lines.ToString()));
                    cmd.Parameters.Add(new SqlParameter("@documentID", documentNo));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read() && rs[0].ToString().Length > 0)
                            total = rs.GetDecimal(0);
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "CheckDropShip"));
            }

            return total;
        }
        public static AddressObject GetShippingInformation(string orderID)
        {
            AddressObject address = null;
            SqlCommand cmd = null;

            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionNavision))
                {
                    dbcon.Open();

                    cmd = new SqlCommand("[dbo].[Applications.NAV.Process.cXML.Documents.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET SHIP TO ADDRESS"));
                    cmd.Parameters.Add(new SqlParameter("@documentID", orderID));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                        if (rs.Read())
                            address = new AddressObject("shipTo"
                                , rs["Ship-To Code"].ToString()
                                , rs["Ship-to Name"].ToString() + "|" + rs["Ship-to Name 2"].ToString()
                                , rs["Ship-to Contact"].ToString()
                                , rs["Ship-to Address"].ToString() + "|" + rs["Ship-to Address 2"].ToString()
                                , rs["Ship-to City"].ToString()
                                , rs["Ship-to County"].ToString()
                                , rs["Ship-to Post Code"].ToString()
                                , rs["Ship-to Country_Region Code"].ToString()
                                , "");
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "GetShippingInformation"));
            }

            return address;
        }
        public static string CheckPurchaseOrder(string orderNumber, string invoiceID)
        {
            if (orderNumber.Length > 0)
            {
                SqlCommand cmd = new SqlCommand();
                try
                {
                    using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionNavision))
                    {
                        dbcon.Open();

                        cmd = new SqlCommand("[dbo].[Applications.NAV.Process.cXML.Documents.Control]", dbcon);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@method", "CHECK FOR PO"));
                        cmd.Parameters.Add(new SqlParameter("@documentID", orderNumber));
                        cmd.Parameters.Add(new SqlParameter("@invoiceNumber", invoiceID));
                        using (SqlDataReader rs = cmd.ExecuteReader())
                            if (rs.Read())
                                return rs[0].ToString();
                    }
                }
                catch (Exception ex)
                {
                    Constants.ERRORS.Add(new Error(ex, cmd, "Prequalifier", "CheckPurchaseOrder"));
                }
            }

            return "";
        }
        #endregion

        #region ECOM_DB
        public static EmailConfig GetEmailConfiguration()
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
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
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "GetEmailConfiguration"));
            }

            string emailUsername = CustomTextEncrypt.Decode(ConfigurationManager.AppSettings["emailUsername"]);
            string password = CustomTextEncrypt.Decode(ConfigurationManager.AppSettings["emailPassword"]);
            string domain = CustomTextEncrypt.Decode(ConfigurationManager.AppSettings["emailDomain"]);

            return new EmailConfig(ConfigurationManager.AppSettings["emailDownHost"], emailUsername, password, domain, ConfigurationManager.AppSettings["emailAdmin"]);
        }
        public static List<EmailFolderConfig> GetEmailFolderConfigs()
        {
            List<EmailFolderConfig> folders = new List<EmailFolderConfig>();
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
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
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "GetEmailFolderConfigs"));
            }

            return folders;
        }
        public static void InsertInvoice(Invoice invoice)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "INSERT LOG"));
                    cmd.Parameters.Add(new SqlParameter("@documentType", "INVOICE"));
                    cmd.Parameters.Add(new SqlParameter("@format", invoice.Format));
                    cmd.Parameters.Add(new SqlParameter("@emailFrom", invoice.EmailFrom));
                    cmd.Parameters.Add(new SqlParameter("@emailSubject", invoice.EmailSubject));
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
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "InsertInvoice"));
            }
        }
        public static string GetVendor(string useragent, string fromid, string secret, string invoiceid, string remitToName)
        {
            string vendor = "";
            SqlCommand cmd = null;

            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
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
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "GetVendor"));
            }

            return vendor;
        }
        public static void GetVendor(ref InvoiceHeader invoice, AddressObject remitTo)
        {
            SqlCommand cmd = null;

            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                    cmd.CommandType = System.Data.CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET VENDOR"));
                    cmd.Parameters.Add(new SqlParameter("@userAgent", invoice.UserAgent));
                    cmd.Parameters.Add(new SqlParameter("@name", remitTo != null ? remitTo.Name : ""));
                    cmd.Parameters.Add(new SqlParameter("@fromIdentity", invoice.FromIdentity));
                    cmd.Parameters.Add(new SqlParameter("@sharedSecret", invoice.SharedSecret));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                        {
                            invoice.Vendor = rs["vendor"].ToString();
                            if (invoice.Vendor == "Invitrogen") invoice.SpecialHandlingAmount = 0.00M;
                            if (invoice.Vendor == "Staples") invoice.SubTotalAmount = invoice.SubTotalAmount - invoice.Tax;
                            if (invoice.Vendor == "Office City" && invoice.OrderDate.Length == 0) invoice.OrderDate = DateTime.Now.ToShortDateString();
                        }
                    }
                }

                if (invoice.Vendor == null)
                    invoice.Vendor = "";
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, "ReadXML", "CheckPONumber"));
            }
        }
        public static void UpdateInvoice(string id, string file, bool retry, bool kwiktagged)
        {
            SqlCommand cmd = null;
            string query = "[dbo].[Ecommerce.ElectronicInvoice.Control]";
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand(query, dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "UPDATE INVOICE"));
                    cmd.Parameters.Add(new SqlParameter("@invoiceID", id));
                    cmd.Parameters.Add(new SqlParameter("@docAlphaPdfFile", file));
                    cmd.Parameters.Add(new SqlParameter("@kwiktagged", kwiktagged ? 1 : 0));
                    cmd.Parameters.Add(new SqlParameter("@invoiceSent", retry ? 0 : 1));
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "UpdateInvoice"));
            }
        }
        public static List<InvoiceHeader> PullInvoices(ref List<InvoiceHeader> invoices, ref List<InvoiceHeader> invoicesOnHold, ref List<InvoiceHeader> emptyVendorInvoices)
        {
            string query = "[dbo].[Ecommerce.ElectronicInvoice.Control]";
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand(query, dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET INVOICES"));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        while (rs.Read())
                        {
                            try
                            {
                                InvoiceHeader invoice = new InvoiceHeader();
                                int parent = GetIntValue(rs, "ID");
                                invoice.UserAgent = GetStringValue(rs, "userAgent");
                                invoice.Vendor = GetStringValue(rs, "vendor");
                                invoice.InvoiceID = GetStringValue(rs, "invoiceID");
                                invoice.Purpose = GetStringValue(rs, "purpose");
                                invoice.InvoiceDate = GetStringValue(rs, "invoiceDate");
                                invoice.OrderID = GetStringValue(rs, "orderId");
                                invoice.OrderDate = GetStringValue(rs, "orderDate");
                                invoice.Tax = GetDecimalValue(rs, "tax");
                                invoice.GrossAmount = GetDecimalValue(rs, "grossAmount");
                                invoice.NetAmount = GetDecimalValue(rs, "netAmount");
                                invoice.DueAmount = GetDecimalValue(rs, "dueAmount");
                                invoice.ShippingAmount = GetDecimalValue(rs, "shippingAmount");
                                invoice.SpecialHandlingAmount = GetDecimalValue(rs, "specialHandlingAmount");
                                invoice.ReceiveDate = GetDateTimeValue(rs, "uploadDate");
                                invoice.PaymentTermNumberOfDays = GetStringValue(rs, "paymentTermDays");
                                invoice.PaymentTermPercentRate = GetStringValue(rs, "paymentTermPercent");
                                invoice.InvoiceDetailDiscount = GetDecimalValue(rs, "invoiceDetailDiscount");
                                invoice.PDFFileName = GetStringValue(rs, "docAlphaPdfFile");
                                int kwiktagged = GetIntValue(rs, "kwiktagged");
                                if (kwiktagged == 1) invoice.Kwiktagged = true;
                                else invoice.Kwiktagged = false;

                                string[] extrinsics = GetStringValue(rs, "extrinsics").Split('|');
                                for (int i = 0; i < extrinsics.Length; i++)
                                {
                                    if (extrinsics[i].Length > 0 && extrinsics[i].Contains(":"))
                                    {
                                        string[] extrs = extrinsics[i].Split(':');
                                        invoice.Extrinsics.Add(new Extrinsic(extrs[0], extrs[1]));
                                    }
                                }

                                decimal shipLineTotal = 0.0M;
                                cmd = new SqlCommand(query, dbcon);
                                cmd.CommandType = CommandType.StoredProcedure;
                                cmd.Parameters.Add(new SqlParameter("@method", "GET INVOICE ROLES AND LINES"));
                                cmd.Parameters.Add(new SqlParameter("@parent", parent));
                                using (SqlDataReader lineRs = cmd.ExecuteReader())
                                {
                                    while (lineRs.Read())
                                    {
                                        invoice.Roles.Add(new AddressObject(GetStringValue(lineRs, "role")
                                            , GetStringValue(lineRs, "addressID")
                                            , GetStringValue(lineRs, "name")
                                            , GetStringValue(lineRs, "deliverTo")
                                            , GetStringValue(lineRs, "street")
                                            , GetStringValue(lineRs, "city")
                                            , GetStringValue(lineRs, "state")
                                            , GetStringValue(lineRs, "postalCode")
                                            , GetStringValue(lineRs, "countryCode")
                                            , GetStringValue(lineRs, "country")));
                                    }

                                    lineRs.NextResult();

                                    while (lineRs.Read())
                                    {
                                        InvoiceLine line = new InvoiceLine();
                                        line.LineNumber = GetIntValue(lineRs, "lineNumber");
                                        line.Quantity = GetDecimalValue(lineRs, "quantity");
                                        line.UnitOfMeasure = GetStringValue(lineRs, "unitOfMeasure");
                                        line.UnitPrice = GetDecimalValue(lineRs, "unitPrice");
                                        line.ReferenceLineNumber = GetIntValue(lineRs, "refrLineNumber");
                                        line.SupplierPartID = GetStringValue(lineRs, "supplierPartID");
                                        line.Description = GetStringValue(lineRs, "description");
                                        line.Tax = GetDecimalValue(lineRs, "tax");
                                        line.LineTotal = GetDecimalValue(lineRs, "lineTotal");
                                        line.ShipLine = GetIntValue(lineRs, "shipLine");

                                        if (line.SupplierPartID.ToUpper().Contains("SHIP")
                                            || line.SupplierPartID.ToUpper().Contains("HAZMAT")
                                            || line.SupplierPartID.ToUpper().Contains("ICE CHARGE")
                                            || line.SupplierPartID.ToUpper().Contains("HANDLING")
                                            || line.SupplierPartID.ToUpper().Contains("TARIFF")
                                            || line.SupplierPartID.ToUpper().Contains("INSURANCE CHARGE")
                                            || (invoice.Vendor == "Praxair" && (line.SupplierPartID.StartsWith("U") || line.SupplierPartID.StartsWith("ZZZ")) && line.SupplierPartID != "ZZZRENT")
                                            || line.ShipLine == 1)
                                        {
                                            shipLineTotal += line.LineTotal;
                                            invoice.ShipLines.Add(line);
                                        }
                                        else
                                        {
                                            InvoiceLine templine = invoice.Lines.Find(l => l.SupplierPartID == line.SupplierPartID && l.UnitPrice == line.UnitPrice && l.ReferenceLineNumber == line.ReferenceLineNumber);

                                            invoice.SubTotalAmount += line.LineTotal;
                                            if (templine == null)
                                                invoice.Lines.Add(line);
                                            else
                                                templine.Quantity += line.Quantity;
                                        }
                                    }
                                }

                                decimal specialhandling = invoice.SpecialHandlingAmount;
                                if (invoice.ShipLines.Find(s => s.UnitPrice == specialhandling) != null)
                                    specialhandling = 0.0M;

                                if (specialhandling == shipLineTotal)
                                    specialhandling = 0.0m;

                                invoice.InvoiceTotal = invoice.SubTotalAmount + invoice.ShippingAmount + invoice.Tax + specialhandling + shipLineTotal;

                                AddressObject shipTo = invoice.Roles.Find(r => r.Role == "shipTo");
                                invoice.ShipType = GetShipType(shipTo);

                                invoice.PO_NAV_Status = CheckPurchaseOrder(invoice.OrderID, invoice.InvoiceID);
                                if (invoice.PO_NAV_Status.StartsWith("DOCID_"))
                                {
                                    invoice.PO_Found = true;
                                    invoice.Document_ID = invoice.PO_NAV_Status.Replace("DOCID_", "");
                                }
                                else
                                    invoice.PO_Found = false;

                                if (invoice.Vendor == "")
                                    emptyVendorInvoices.Add(invoice);
                                else if (Constants.SkipInvoice(ref invoice))
                                    invoices.Add(invoice);
                                else
                                    invoicesOnHold.Add(invoice);
                            }
                            catch (Exception ex)
                            {
                                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "PullInvoices"));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "PullInvoices"));
            }

            return invoices;
        }
        public static string GetShipType(AddressObject shipTo)
        {
            string dropShip = "DROP_SHIP", query = "[dbo].[Ecommerce.ElectronicInvoice.Control]";
            SqlCommand cmd = null;

            if (shipTo != null)
            {
                try
                {
                    using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                    {
                        dbcon.Open();
                        cmd = new SqlCommand(query, dbcon);
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add(new SqlParameter("@method", "GET SHIPTYPE"));
                        cmd.Parameters.Add(new SqlParameter("@name", shipTo.Name));
                        cmd.Parameters.Add(new SqlParameter("@street", shipTo.Street));
                        cmd.Parameters.Add(new SqlParameter("@deliverTo", shipTo.DeliverTo));
                        cmd.Parameters.Add(new SqlParameter("@city", shipTo.City));
                        cmd.Parameters.Add(new SqlParameter("@state", shipTo.State));
                        cmd.Parameters.Add(new SqlParameter("@postalCode", shipTo.PostalCode));
                        using (SqlDataReader rs = cmd.ExecuteReader())
                        {
                            if (rs.Read())
                                dropShip = rs["shipType"].ToString();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Constants.ERRORS.Add(new Error(ex, cmd, "Database", "CheckDropShip"));
                    dropShip = Constants.CheckDropShip(shipTo);
                }
            }

            return dropShip;
        }
        public static string CheckInvoice(string invoiceID, string orderID)
        {
            SqlCommand cmd = null;
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "CHECK INVOICE"));
                    cmd.Parameters.Add(new SqlParameter("@invoiceID", invoiceID));
                    cmd.Parameters.Add(new SqlParameter("@orderID", orderID));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                        {
                            if (rs["invoiceSent"].ToString() == "1")
                                return rs["docAlphaDate"].ToString();
                            else
                                return rs["invoiceSent"].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "CheckInvoice"));
            }
            return "NOT_FOUND";
        }
        public static string GetGlAccountNumber(string shiptype, string supplierPartNo, string description)
        {
            string glaccount = "", query = "[dbo].[Ecommerce.ElectronicInvoice.Control]";
            SqlCommand cmd = null;

            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand(query, dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "GET GL ACCOUNT"));
                    cmd.Parameters.Add(new SqlParameter("@shipType", shiptype));
                    cmd.Parameters.Add(new SqlParameter("@supplierPartId", supplierPartNo));
                    cmd.Parameters.Add(new SqlParameter("@description", description));
                    using (SqlDataReader rs = cmd.ExecuteReader())
                    {
                        if (rs.Read())
                            glaccount = rs["glAccountNo"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "CheckDropShip"));
            }

            return glaccount;
        }
        public static int InsertBatchInformation(string shipType, int batchPoNotFound, string batch, string vendor)
        {
            int id = 0;
            SqlCommand cmd = null;

            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    cmd = new SqlCommand("[dbo].[Ecommerce.ElectronicInvoice.Control]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@method", "INSERT BATCH INFO"));
                    cmd.Parameters.Add(new SqlParameter("@shipType", shipType));
                    cmd.Parameters.Add(new SqlParameter("@batch", batch));
                    cmd.Parameters.Add(new SqlParameter("@vendor", vendor));
                    cmd.Parameters.Add(new SqlParameter("@batchPoNotFound", batchPoNotFound));
                    using (SqlDataReader reader = cmd.ExecuteReader())
                        if (reader.Read())
                            id = GetIntValue(reader, 0);
                }
            }
            catch (Exception ex)
            {
                Constants.ERRORS.Add(new Error(ex, cmd, "Database", "InsertBatchInformation"));
            }
            return id;
        }
        #endregion

        #region Database Functions
        private static string GetStringValue(SqlDataReader rs, string column)
        {
            try { return rs[column].ToString(); }
            catch { return ""; }
        }
        private static int GetIntValue(SqlDataReader rs, string column)
        {
            try { return int.Parse(rs[column].ToString()); }
            catch { return 0; }
        }
        private static int GetIntValue(SqlDataReader rs, int column)
        {
            try { return int.Parse(rs[column].ToString()); }
            catch { return 0; }
        }
        private static decimal GetDecimalValue(SqlDataReader rs, string column)
        {
            try { return decimal.Parse(rs[column].ToString()); }
            catch { return 0.00M; }
        }
        private static DateTime GetDateTimeValue(SqlDataReader rs, string column)
        {
            try { return DateTime.Parse(rs[column].ToString()); }
            catch { return new DateTime(1753, 1, 1); }
        }
        #endregion

    }
}