using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
{
    public class EmailConfig
    {
        public EmailConfig(string host, string username, string password, string domain, string admin)
        {
            Host = host;
            Username = username;
            Password = password;
            Domain = domain;
            AdminEmail = admin;
        }

        public string Host { get; }
        public string Username { get; }
        public string Password { get; }
        public string Domain { get; }
        public string AdminEmail { get; }
    }
}
