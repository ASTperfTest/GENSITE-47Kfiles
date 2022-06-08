using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SQLite;
using System.Text;
using System.Data;
using System.Data.SqlClient;

public partial class grade : System.Web.UI.Page
{
    private string currentFolderPath = System.Configuration.ConfigurationManager.AppSettings.GetValues("Path")[0];
    private string dbpatch = System.Configuration.ConfigurationSettings.AppSettings["DB"].ToString();
    private int type = int.Parse(System.Configuration.ConfigurationSettings.AppSettings["Type"].ToString());
	protected string title = System.Configuration.ConfigurationSettings.AppSettings["TypeName"].ToString();
    
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            BindDataList();
        } 
    }

    protected void BindDataList()
    {
        SQLiteConnection cnn = new SQLiteConnection("Data Source=" + dbpatch);
        DataTable dt = new DataTable();
        try
        {
            cnn.Open();
            SQLiteCommand mycommand = new SQLiteCommand(cnn);
            SQLiteParameter activeType = new SQLiteParameter("@TYPE");
            mycommand.Parameters.Add(activeType);
            mycommand.Parameters["@TYPE"].Value = type;
            mycommand.CommandText = "select UID,CREATOR,CREATOR_DISPLAY_NAME,IMAGE_PATH,COMPLETED,SORT_ORDER from LOGO Where TYPE = @TYPE";
            SQLiteDataReader reader = mycommand.ExecuteReader();
            dt.Load(reader);
            cnn.Close();
        }
        catch
        {

        }		 

        DataList1.DataSource = dt;
        DataList1.DataKeyField = "UID";
        DataList1.DataBind();
    }

    protected void DataList1_EditCommand(object source, DataListCommandEventArgs e)
    {
        DataList1.EditItemIndex = e.Item.ItemIndex;
        BindDataList();
    }

    protected void DataList1_CancelCommand(object source, DataListCommandEventArgs e)
    {
        DataList1.EditItemIndex = -1;
        BindDataList();
    }

    protected void DataList1_UpdateCommand(object source, DataListCommandEventArgs e)
    {
        SQLiteConnection cnn = new SQLiteConnection("Data Source=" + dbpatch);
        DataTable dt = new DataTable();
        try
        {
            cnn.Open();
            SQLiteCommand mycommand = new SQLiteCommand(cnn);
            SQLiteParameter date = new SQLiteParameter("@DATETIME");
            SQLiteParameter sort_order = new SQLiteParameter("@SORT_ORDER");
            SQLiteParameter dbuid = new SQLiteParameter("@UID");
            SQLiteParameter dcoComplete = new SQLiteParameter("@COMPLETED");
            SQLiteParameter editor = new SQLiteParameter("@EDITOR");

            mycommand.Parameters.Add(date);
            mycommand.Parameters["@DATETIME"].Value = DateTime.Now.ToString();
            mycommand.Parameters.Add(sort_order);
			int rankorder = int.Parse(((TextBox)e.Item.FindControl("TextRank")).Text);
            mycommand.Parameters["@SORT_ORDER"].Value = rankorder;
            mycommand.Parameters.Add(dbuid);
            mycommand.Parameters["@UID"].Value = Convert.ToInt32(DataList1.DataKeys[e.Item.ItemIndex].ToString());
            mycommand.Parameters.Add(dcoComplete);
            mycommand.Parameters["@COMPLETED"].Value = ((CheckBox)e.Item.FindControl("CheckBoxEdit")).Checked;
            mycommand.Parameters.Add(editor);
            mycommand.Parameters["@EDITOR"].Value = "admin";

            mycommand.CommandText = "update LOGO Set LAST_MODIFIER =@EDITOR,LAST_MODIFY_DATETIME = @DATETIME,COMPLETED=@COMPLETED,SORT_ORDER=@SORT_ORDER where UID =@UID";
            SQLiteDataReader reader = mycommand.ExecuteReader();
            dt.Load(reader);
            cnn.Close();
        }
        catch (Exception ex)
        {

        }
        DataList1.EditItemIndex = -1;
        BindDataList();
    }

}
