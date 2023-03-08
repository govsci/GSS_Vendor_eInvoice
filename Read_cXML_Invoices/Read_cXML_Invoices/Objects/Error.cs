using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace Read_cXML_Invoices.Objects
{
    public class Error
    {
        public Error(Exception ex, string cl, string function)
        {
            Ex = ex;
            Class = cl;
            Function = function;
        }

        public Error(Exception ex, SqlCommand cmd, string cl, string function)
        {
            Ex = ex;
            Cmd = cmd;
            Class = cl;
            Function = function;
        }

        public Exception Ex { get; }
        public SqlCommand Cmd { get; }
        public string Class { get; }
        public string Function { get; }
    }
}
