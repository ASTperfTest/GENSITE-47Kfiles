using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.SqlClient;

public class FarmDB
{
    public static string connString
    {
        get
        {
            SqlConnectionStringBuilder connStr = new SqlConnectionStringBuilder();
            connStr.DataSource = Farm2009Config.DB_SOURCENAME;
            connStr.InitialCatalog = Farm2009Config.DB_NAME;
            connStr.UserID = Farm2009Config.DB_USER;
            connStr.Password = Farm2009Config.DB_PW;
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

        conn.ConnectionString = FarmDB.connString;
        return conn;
    }

}
