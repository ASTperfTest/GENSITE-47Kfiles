using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
public partial class Edit_Class_node : System.Web.UI.Page
{
    public Setting dbconfig = new Setting();
    public string class_id;
    public string parent_id;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["id"] != null)
        {
            class_id = Request.QueryString["id"].ToString();
            parent_id = Request.QueryString["id"].ToString();
        }

        SqlDataSource1.ConnectionString = dbconfig.ConnectionSettings();
        SqlDataSource1.SelectCommand = "select * from type where datalevel = 2 and dataparent ='"+class_id+"' order by sortvalue asc";
        
        if (!IsPostBack)
        {
            get_f_class();

        }



    }
    protected void SqlDataSource1_Selecting(object sender, SqlDataSourceSelectingEventArgs e)
    {

    }

     protected void get_f_class()
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "select * from type where classid=@Classid";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classid", SqlDbType.Int).Value = Convert.ToInt32(class_id);
        SqlDataReader reader = cmd.ExecuteReader();
        if (reader.Read())
        {
            Label_Class.Text = "父分類名稱：" + reader["classname"].ToString();
        
        }
        reader.Close();
        conn.Close();
    }

    protected void GridView1_DataBound(object sender, EventArgs e)
    {
        for (int i = 0; i <= GridView1.Rows.Count - 1; i++)
        {
            HyperLink link = new HyperLink();
            link.NavigateUrl = "class_node_edit.aspx?id=" +GridView1.Rows[i].Cells[0].Text+"&&parentid="+parent_id+"";
            link.Text = "管理";
             
            GridView1.Rows[i].Cells[2].Text = "";
            GridView1.Rows[i].Cells[2].Controls.Add(link);
        }
    }
}
