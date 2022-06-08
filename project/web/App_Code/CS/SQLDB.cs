using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.SqlClient;

public class SQLDB
{
    public static string connString
    {
        get
        {
            SqlConnectionStringBuilder connStr = new SqlConnectionStringBuilder();
            connStr.DataSource = ".";
            connStr.InitialCatalog = "FishBowl";
            connStr.UserID = "sa";
            connStr.Password = "gss";
            connStr.ConnectTimeout = 20;
            connStr.PersistSecurityInfo = true;
            connStr.Pooling = true;
            connStr.MaxPoolSize = 250;
            connStr.MinPoolSize = 5;

            return connStr.ConnectionString;
        }
    }

    public static SqlConnection createConnection()
    {
        SqlConnection conn = new SqlConnection();

        conn.ConnectionString = SQLDB.connString;
        conn.Open();
        return conn;
    }

}
