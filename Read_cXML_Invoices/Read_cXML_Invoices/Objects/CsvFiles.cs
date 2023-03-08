using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Read_cXML_Invoices.Objects
{
    public class CsvFiles
    {
        public CsvFiles(string sheetName, string content)
        {
            SheetName = sheetName;
            Content = content;
        }

        public string SheetName { get; }
        public string Content { get; }
    }
}
