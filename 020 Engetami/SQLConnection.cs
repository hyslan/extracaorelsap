using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;

namespace SQLConnection
{
    public class SqlConnectionBdMlg
    {
        public SqlConnection BD_MLG_Query()
        {
            string connectionString = "Data Source=10.66.42.188;Initial Catalog=BD_MLG;Integrated Security=SSPI;";
            SqlConnection connect = new SqlConnection(connectionString);
            return connect;
        }
    }
}
