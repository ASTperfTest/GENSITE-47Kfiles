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

public partial class Edit_class_edit : System.Web.UI.Page
{
    public Setting dbconfig = new Setting();
    public string class_id;
    protected void Page_Load(object sender, EventArgs e)
    {
       
            if (Request.QueryString["id"].ToString() == "0")
            {
                class_add();

            }
            else
            {
                class_id = Request.QueryString["id"].ToString();
               
            
            }
            if (!IsPostBack && Request.QueryString["id"].ToString() != "0")
            {
                class_edit();
            
            }
       
    }

    protected void btn_Add_Click(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "insert into type(classname,datalevel,sortvalue) values(@Classname,1,@Sortvalue)";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classname", SqlDbType.VarChar).Value = Txt_Class.Text;
        cmd.Parameters.Add("@Sortvalue", SqlDbType.Int).Value = Convert.ToInt32(Txt_sortvalue.Text);
        cmd.ExecuteNonQuery();
        conn.Close();
        clear();
        Response.Redirect("class.aspx");
    }

    protected void btn_Edit_Click(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "update type set classname =@Classname,sortvalue=@Sortvalue where classid=@Classid";
        Response.Write(strSQL);
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classname", SqlDbType.VarChar).Value = Txt_Class.Text;
        cmd.Parameters.Add("@Classid", SqlDbType.Int).Value = Convert.ToInt32(class_id);
        cmd.Parameters.Add("@Sortvalue", SqlDbType.Int).Value = Convert.ToInt32(Txt_sortvalue.Text);
        cmd.ExecuteNonQuery();
        conn.Close();
        clear();
        Response.Redirect("class.aspx");
    }

    protected void btn_Cancel_Click(object sender, EventArgs e)
    {
        clear();
        Response.Redirect("class.aspx");
    }

    protected void btn_Delete_Click(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "delete from type where classid=@Classid";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classid", SqlDbType.Int).Value = Convert.ToInt32(class_id);
        cmd.ExecuteNonQuery();
        conn.Close();
        clear();
        Response.Redirect("class.aspx");
    }

    protected void class_add()
    {
        Label1.Visible = true;
        Txt_Class.Visible = true;
        btn_Add.Visible = true;
        btn_Cancel.Visible = true;
        Sort.Visible = true;
        Txt_sortvalue.Visible = true;

    }

    protected void class_edit()
    {
        Label1.Visible = true;
        Txt_Class.Visible = true;
        btn_Edit.Visible = true;
        btn_Cancel.Visible = true;
        btn_Delete.Visible = true;
        Sort.Visible = true;
        Txt_sortvalue.Visible = true;

        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "select * from type where 1=1 and datalevel = 1 and classid=@Classid";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classid", SqlDbType.Int).Value = Convert.ToInt32(class_id);
        SqlDataReader reader = cmd.ExecuteReader();
        if (reader.Read())
        {
            Txt_Class.Text = reader["classname"].ToString();
            Txt_sortvalue.Text = reader["sortvalue"].ToString();
        }
        reader.Close();

        strSQL = "select count(*) as total from type where 1=1 and datalevel = 2 and dataparent=@Classid";
        SqlCommand cmd2 = new SqlCommand(strSQL, conn);
        cmd2.Parameters.Add("@Classid", SqlDbType.Int).Value = Convert.ToInt32(class_id);
        SqlDataReader reader2 = cmd2.ExecuteReader();
        if (reader2.Read())
        {
            int x = Convert.ToInt32(reader2["total"]);
            if (x > 0)
            {
                btn_Delete.Enabled = false;
                lbl_worng.Visible = true;
            }

        }
        reader2.Close();
        conn.Close();
    }

    protected void clear()
    {
        Label1.Visible = false;
        Txt_Class.Text = "";
        Txt_Class.Visible = false;
        btn_Add.Visible = false;
        btn_Edit.Visible = false;
        btn_Delete.Visible = false;
        btn_Cancel.Visible = false;
        lbl_worng.Visible = false;

    }

}
