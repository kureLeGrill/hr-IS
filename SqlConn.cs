using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace LEGAL
{
    class SqlConn
    {
        public SqlConnection sqlConnection;

        public SqlConnection OpenSqlConn(string connString)
        {
            sqlConnection = new SqlConnection();
            sqlConnection.ConnectionString = connString;
            sqlConnection.Open();
            return sqlConnection;
        }
        public void CloseSqlConn(SqlConnection connection)
        {
            connection.Close();
        }
    }
}
