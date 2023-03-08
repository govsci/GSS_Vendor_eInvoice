using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Read_cXML_Invoices.Objects;
using System.Net;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Read_cXML_Invoices.Classes
{
    public class Email
    {
        public static void SendEmail(string msg, string subject, string emailFrom, string emailTo, string emailCC, string emailBCC, string file, bool html)
        {
            EmailConfig emailConfig = GetEmailConfig();
            if (emailTo.Length == 0)
                emailTo = emailConfig.AdminEmail;

            MailMessage mail = new MailMessage();
            mail.IsBodyHtml = html;
            string error = "";

            //Message (Body)
            if (msg.Length == 0)
                error += "\nBody of the email is empty";
            else
                mail.Body = msg;

            //Subject
            if (subject.Length == 0)
                error += "\nSubject of the email is blank";
            else
                mail.Subject = subject;

            //CC
            if (emailCC.Length > 0)
            {
                if (emailCC.Contains(';'))
                {
                    string[] emails = emailCC.Split(';');
                    foreach (string email in emails)
                    {
                        if (email.Length > 0)
                        {
                            if (TestEmail(email, false))
                                mail.CC.Add(new MailAddress(email));
                            else
                                error += "\nEmail Carbon Copy (CC) Address is not valid: " + email;
                        }
                    }
                }
                else
                {
                    if (TestEmail(emailCC, false))
                        mail.CC.Add(new MailAddress(emailCC));
                    else
                        error += "\nEmail Carbon Copy (CC) Address is not valid: " + emailCC;
                }
            }

            //From
            if (emailFrom.Length == 0)
                mail.From = new MailAddress("ecommercesystem@govsci.com");
            else
            {
                if (emailFrom.Contains(";"))
                    error += "\nEmail From Address is invalid: " + emailFrom;
                else
                {
                    if (TestEmail(emailFrom, false))
                        mail.From = new MailAddress(emailFrom);
                    else
                        error += "\nEmail From Address is invalid: " + emailFrom;
                }
            }

            //To
            if (emailTo.Contains(';'))
            {
                string[] emails = emailTo.Split(';');
                for (int i = 0; i < emails.Length; i++)
                {
                    if (TestEmail(emails[i], true))
                    {
                        mail.To.Add(new MailAddress(emails[i]));
                    }
                    else
                        error += "\nEmail To Address is not valid: " + emails[i];
                }
            }
            else
            {
                if (TestEmail(emailTo, true))
                    mail.To.Add(new MailAddress(emailTo));
                else
                    error += "\nEmail To Address is not valid: " + emailTo;
            }

            //BCC
            if (emailBCC.Length > 0)
            {
                if (emailBCC.Contains(';'))
                {
                    string[] emails = emailBCC.Split(';');
                    foreach (string email in emails)
                    {
                        if (email.Length > 0)
                        {
                            if (TestEmail(email, false))
                                mail.Bcc.Add(new MailAddress(email));
                            else
                                error += "\nEmail Blind Carbon Copy (BCC) Address is not valid: " + email;
                        }
                    }
                }
                else
                {
                    if (TestEmail(emailBCC, false))
                        mail.Bcc.Add(new MailAddress(emailBCC));
                    else
                        error += "\nEmail Blind Carbon Copy (BCC) Address is not valid: " + emailBCC;
                }
            }

            //File
            if (file.Length > 0)
            {
                if (file.Contains(';'))
                {
                    string[] files = file.Split(';');
                    foreach (string fil in files)
                        if (fil.Length > 0 && File.Exists(fil))
                            mail.Attachments.Add(new Attachment(fil));
                }
                else if (file.Contains(','))
                {
                    string[] files = file.Split(',');
                    foreach (string fil in files)
                        if (fil.Length > 0 && File.Exists(fil))
                            mail.Attachments.Add(new Attachment(fil));
                }
                else if (File.Exists(file))
                    mail.Attachments.Add(new Attachment(file));
            }

            if (error.Length == 0)
                Send(mail, emailConfig);
            else
                throw new Exception("The following errors have occurred: " + error);
        }
        public static void SendErrorMessage()
        {
            string folder = Constants.ReportPath + DateTime.Now.ToString(@"yyyy\\MM\\dd\\");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            string file = DateTime.Now.ToString("yyyyMMddhhmmssfffffff") + ".ErrorReports.html";

            string msg = "Team Foundation Server: Internal Apps > Zach Projects > Read_cXML_Invoices<br><br>";
            msg += "Please check the error log on the server MGT-SCH, " + folder + file;

            string errormsg = "<table border='1'><tbody><tr><th>Class</th><th>Function</th><th>SQL CMD</th><th>Error</th></tr>";

            foreach (Error err in Constants.ERRORS)
            {
                string row = $"<tr><td>{err.Class}</td><td>{err.Function}</td>";
                
                if (err.Cmd != null)
                {
                    string query = err.Cmd.CommandText;
                    foreach (SqlParameter para in err.Cmd.Parameters)
                        query += " " + para.ParameterName + "='" + para.Value + "', ";
                    row += $"<td>{query}</td>";
                }
                else
                    row += $"<td></td>";

                row += $"<td>{err.Ex.ToString()}</td></tr>";
                errormsg += row;
            }

            errormsg += "</tbody></table>";

            File.WriteAllText(folder + file, errormsg);

            SendEmail(msg, "Read_cXML_Invoices Error", "", "zlingelbach@govsci.com", "", "", "", true);
        }
        private static void Send(MailMessage mail, EmailConfig emailConfig)
        {
            try
            {
                SmtpClient client = new SmtpClient(emailConfig.Host);
                client.Credentials = new NetworkCredential(emailConfig.Username, emailConfig.Password, emailConfig.Domain);
                client.Send(mail);
            }
            catch (Exception)
            {

            }
        }
        private static EmailConfig GetEmailConfig()
        {
            try
            {
                using (SqlConnection dbcon = new SqlConnection(Constants.DbConnectionEcommerce))
                {
                    dbcon.Open();
                    SqlCommand cmd = new SqlCommand("[dbo].[Ecommerce.Get.Email.Configuration]", dbcon);
                    cmd.CommandType = CommandType.StoredProcedure;
                    using (SqlDataReader rs = cmd.ExecuteReader())
                        if (rs.Read())
                            return new EmailConfig(rs["host"].ToString(), rs["username"].ToString(), rs["password"].ToString(), rs["domain"].ToString(), rs["admin"].ToString());
                }
            }
            catch (Exception)
            {
            }

            return new EmailConfig("webmail.govsci.com", "ecommercesystem", "Secure1", "GSS1", "zlingelbach@govsci.com");
        }
        private static bool TestEmail(string email, bool req)
        {
            try
            {
                if (email.Length > 0)
                    new MailAddress(email);
                else if (req)
                    return false;
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
