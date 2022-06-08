using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.SqlClient;

/// <summary>
/// TreasureDB 的摘要描述
/// </summary>
public class TreasureHuntDB
{
    public static string connString
    {
        get
        {
            SqlConnectionStringBuilder connStr = new SqlConnectionStringBuilder();
            connStr.DataSource = TreasureHuntConfig.DB_SOURCENAME;
            connStr.InitialCatalog = TreasureHuntConfig.DB_NAME;
            connStr.UserID = TreasureHuntConfig.DB_USER;
            connStr.Password = TreasureHuntConfig.DB_PW;
            connStr.ConnectTimeout = 20;
            connStr.PersistSecurityInfo = true;
            connStr.Pooling = true;
            connStr.MaxPoolSize = 250;
            connStr.MinPoolSize = 5;

            return connStr.ConnectionString;
        }
    }

    public static string coaConnString
    {
        get
        {
            SqlConnectionStringBuilder connStr = new SqlConnectionStringBuilder();
            connStr.DataSource = TreasureHuntConfig.COA_DB_SOURCENAME;
            connStr.InitialCatalog = TreasureHuntConfig.COA_DB_NAME;
            connStr.UserID = TreasureHuntConfig.COA_DB_USER;
            connStr.Password = TreasureHuntConfig.COA_DB_PW;
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

        conn.ConnectionString = TreasureHuntDB.connString;
        return conn;
    }

    public static SqlConnection createCoaConnection()
    {
        SqlConnection conn = new SqlConnection();

        conn.ConnectionString = TreasureHuntDB.coaConnString;
        return conn;
    }
}
