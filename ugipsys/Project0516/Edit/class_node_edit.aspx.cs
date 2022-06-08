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

public partial class Edit_class_node_edit : System.Web.UI.Page
{
    public Setting dbconfig = new Setting();
    public string class_id;
    public string parent_id;
    public string parent;
    protected void Page_Load(object sender, EventArgs e)
    {

        if (Request.QueryString["id"].ToString() == "0")
        {
            class_add();

        }
        else
        {
            class_id = Request.QueryString["id"].ToString();
            parent_id = Request.QueryString["parentid"].ToString();

        }
        if (!IsPostBack && Request.QueryString["id"].ToString() != "0")
        {
            get_ddl();
            class_edit();
            

        }
        else if(!IsPostBack && Request.QueryString["id"].ToString() == "0")
        {
            get_ddl();
        }

    }

    protected void btn_Add_Click(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "insert into type(classname,datalevel,dataparent,sortvalue) values(@Classname,2,@Dataparent,@Sortvalue)";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classname", SqlDbType.VarChar).Value = Txt_Class.Text;
        cmd.Parameters.Add("@Dataparent", SqlDbType.Int).Value = ddl_class1.SelectedValue;
        cmd.Parameters.Add("@Sortvalue", SqlDbType.Int).Value = Convert.ToInt32(Txt_sortvalue.Text);
        cmd.ExecuteNonQuery();
        conn.Close();
        clear();
        Response.Redirect("class_node.aspx?id=" + ddl_class1.SelectedValue);
    }

    protected void btn_Edit_Click(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "update type set classname =@Classname,dataparent=@Dataparent,sortvalue=@Sortvalue where classid=@Classid and dataparent=@Parentid";
       
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classname", SqlDbType.VarChar).Value = Txt_Class.Text;
        cmd.Parameters.Add("@Dataparent", SqlDbType.Int).Value = ddl_class1.SelectedValue;
        cmd.Parameters.Add("@Classid", SqlDbType.Int).Value = Convert.ToInt32(class_id);
        cmd.Parameters.Add("@Parentid", SqlDbType.Int).Value = Convert.ToInt32(parent_id);
        cmd.Parameters.Add("@Sortvalue", SqlDbType.Int).Value = Convert.ToInt32(Txt_sortvalue.Text);
        cmd.ExecuteNonQuery();
        conn.Close();
        clear();
        Response.Redirect("class_node.aspx?id=" + ddl_class1.SelectedValue);
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
        Response.Redirect("class_node.aspx?id=" + ddl_class1.SelectedValue);
    }

    protected void class_add()
    {
        Label1.Visible = true;
        Txt_Class.Visible = true;
        btn_Add.Visible = true;
        btn_Cancel.Visible = true;
        Label2.Visible = true;
        ddl_class1.Visible = true;
      
       

    }

    protected void class_edit()
    {
        Label1.Visible = true;
        Txt_Class.Visible = true;
        btn_Edit.Visible = true;
        btn_Cancel.Visible = true;
        btn_Delete.Visible = true;
        Label2.Visible = true;
        ddl_class1.Visible = true;
        
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "select * from type where 1=1 and datalevel = 2 and classid=@Classid";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Classid", SqlDbType.Int).Value = Convert.ToInt32(class_id);
        
        SqlDataReader reader = cmd.ExecuteReader();
        if (reader.Read())
        {
            Txt_Class.Text = reader["classname"].ToString();
            parent = reader["dataparent"].ToString();
            Txt_sortvalue.Text = reader["sortvalue"].ToString();
        }
        reader.Close();
       

        for (int i = 0; i < ddl_class1.Items.Count; i++)
        {
           
            if (parent == ddl_class1.Items[i].Value)
            {
                
                ddl_class1.Items[i].Selected = true;
            }
        }
        //strSQL = "select count(*) as total from class where 1=1 and datalevel = 2 and dataparent=" + class_id;
        //SqlCommand cmd2 = new SqlCommand(strSQL, conn);
        //SqlDataReader reader2 = cmd2.ExecuteReader();
        //if (reader2.Read())
        //{
        //    int x = Convert.ToInt32(reader2["total"]);
        //    if (x > 0)
        //    {
        //        btn_Delete.Enabled = false;
        //        lbl_worng.Visible = true;
        //    }

        //}
        //reader2.Close();
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
    protected void get_ddl()
    {
        ddl_class1.Items.Clear();
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "select * from type where datalevel = 1";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        
        SqlDataReader reader = cmd.ExecuteReader();
        ddl_class1.Items.Add("--請選擇--");
        while (reader.Read())
        {

            ListItem li = new ListItem(reader["classname"].ToString(), reader["classid"].ToString());
            ddl_class1.Items.Add(li);

        }

    
    }


}
