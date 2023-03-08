using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
{
    public class EmailFolderConfig
    {
        public EmailFolderConfig(string documentType, string emailFolder, string localFolder)
        {
            DocumentType = documentType;
            EmailFolder = emailFolder;
            LocalFolder = localFolder;
        }

        public string DocumentType { get; }
        public string EmailFolder { get; }
        public string LocalFolder { get; }
    }
}
