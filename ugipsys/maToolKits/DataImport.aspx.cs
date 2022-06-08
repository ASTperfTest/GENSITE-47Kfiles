using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Data;
using GSS.Vitals.COA.Data;

public partial class maToolKits_DataImport : System.Web.UI.Page
{
    private string iCTUnit, MemberID;
    string Script = "<script>alert('連線逾時或尚未登入，請重新登入');window.close();</script>";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
        else
        {

            iCTUnit = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitEvent"].ToString();
            MemberID = Session["userID"].ToString();
        }
    }

    protected void btnImport_Click(object sender, EventArgs e)
    {

        ProcessTransData();

    }
    protected void btnClearDB_Click(object sender, EventArgs e) 
    {
        string strQueryiCUItemScript = @"SELECT gicuitem FROM HistoryList";
        string strDeleteScript = @"DELETE FROM CuDTGeneric WHERE iCUItem = @iCUItem";

        using (var reader = SqlHelper.ReturnReader("ConnString", strQueryiCUItemScript))
        {
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    SqlHelper.ExecuteNonQuery("ConnString", strDeleteScript,
                        DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", reader["gicuitem"].ToString()));
                }
            }
        }
        string strTruncateScript = @"TRUNCATE TABLE HistoryList";
        SqlHelper.ExecuteNonQuery("ConnString", strTruncateScript);
    }


    private void ProcessTransData()
    {
        string f = this.fileCSV.FileName;
        FileInfo fi = new FileInfo(f);
        string name = fileCSV.FileName;
        string savePath = @"~\public\DataImport\" + name;
        //將檔案先上傳到server上
        fileCSV.SaveAs(Server.MapPath(savePath));
        DataTable dt = this.GetCSVData(@"~\public\DataImport\", fi.Name);

        string strInsertScript = @"INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, iEditor, iDept, xBody) 
                                  VALUES (@iBaseDSD, @iCTUnit, @fCTUPublic, @iEditor, @iDept, @xBody) ";
        foreach (DataRow RowItem in dt.Rows)
        {
            // 寫入CuDTGeneric
            SqlHelper.ExecuteNonQuery("ConnString", strInsertScript,
                DbProviderFactories.CreateParameter("ConnString", "@iBaseDSD", "@iBaseDSD", "47"),
                DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnit),
                DbProviderFactories.CreateParameter("ConnString", "@fCTUPublic", "@fCTUPublic", "Y"),
                DbProviderFactories.CreateParameter("ConnString", "@iEditor", "@iEditor", MemberID),
                DbProviderFactories.CreateParameter("ConnString", "@iDept", "@iDept", "0"),
                DbProviderFactories.CreateParameter("ConnString", "@xBody", "@xBody", RowItem[3].ToString()));
            // 取得剛剛寫入的iCUItem
            string strQueryScript = @"SELECT TOP 1 iCUItem FROM CuDTGeneric WHERE iCTUnit = @iCTUnit ORDER BY iCUItem DESC";
            using (var reader = SqlHelper.ReturnReader("ConnString", strQueryScript,
                DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnit)))
            {
                if (reader.HasRows)
                {
                    // 寫入HistoryList
                    if (reader.Read())
                    {
                        string strHistoryListScript = @"INSERT INTO HistoryList (gicuitem, Year, Month, Day) 
                                                VALUES (@gicuitem, @Year, @Month, @Day)";
                        SqlHelper.ExecuteNonQuery("ConnString", strHistoryListScript,
                            DbProviderFactories.CreateParameter("ConnString", "@gicuitem", "@gicuitem", reader["iCUItem"].ToString()),
                            DbProviderFactories.CreateParameter("ConnString", "@Year", "@Year", RowItem[0]),
                            DbProviderFactories.CreateParameter("ConnString", "@Month", "@Month", RowItem[1]),
                            DbProviderFactories.CreateParameter("ConnString", "@Day", "@Day", RowItem[2]));
                    }
                }
            }
        }
    }

    public DataTable GetCSVData(string savePath, string sheetname)
    {
        System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Text;'", Server.MapPath(savePath)));
        System.Data.OleDb.OleDbDataAdapter adt = new System.Data.OleDb.OleDbDataAdapter("select * from [" + sheetname + "]", conn);
        DataSet ds = new DataSet();
        adt.Fill(ds);
        DataTable dt = ds.Tables[0];
        return dt;
    }





}