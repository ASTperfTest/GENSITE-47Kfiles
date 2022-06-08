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
using System.IO;
using System.Data.SqlClient;

public partial class Step1_2 : System.Web.UI.Page
{
    public string id_no;
    public string nums;
    public string user_idno;
    public string choice;
    public Setting dbconfig = new Setting();
   
    protected void Page_Load(object sender, EventArgs e)
    {
        id_no = Session["User_id"].ToString();
        user_idno = Session["Name"].ToString();
     	  string path = Session["PublicPath"].ToString();
        if (!IsPostBack) {
            get_data(path);
        }
    }

    //下一步
    protected void next_Click(object sender, EventArgs e)
    {
        if (Txt_URL.Text == "") {
            insert_first(user_idno);
            insert_second(user_idno);
            next.Visible = true;            
            delete.Visible = true;
            Session.Add("Genname", Txt_Name.Text);
            if (Session["check"] == null) {                
                Response.Redirect("./step2.aspx");
            }
            else {                
                Response.Redirect("./Step2_2.aspx"); 
            }            
        }
        else {
            insert_first(user_idno);
            insert_second(user_idno);
            next.Visible = false;
            save.Visible = false;            
            Session.Add("URL", "y");
            Response.Redirect("./Finish.aspx");
        }
    }

    //儲存設定
    protected void save_Click(object sender, EventArgs e)
    {
      if (Txt_URL.Text == "") {
        insert_first(user_idno);
        insert_second(user_idno);
        clear();
        get_data(Session["PublicPath"].ToString());
      }
      else {
        insert_first(user_idno);
        insert_second(user_idno);
        clear();
        get_data(Session["PublicPath"].ToString());
      }
    }

    //刪除主題館
    protected void delete_Click(object sender, EventArgs e)
    {      
      string test = Session["User_id"].ToString();
      string delSQL = "DELETE FROM CatTreeRoot WHERE CtRootID = '" + test + "'";
      string delSQL2 = "DELETE FROM NodeInfo WHERE CtRootID = '" + test + "'";
      dbconfig.deletetitle(delSQL, delSQL2, test, Server.MapPath(dbconfig.Filepath()), Session["GipDsdPath"].ToString());
      Response.Redirect("./index.aspx");
    }
  
    //設定下拉
    protected void SetDDL(DropDownList DDL, string type1)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        string strSQL = "Select * from type where 1 = 1 AND datalevel = 1";
        conn.Open();
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        SqlDataReader get_DDL = cmd.ExecuteReader();
        DDL.Items.Add("--請選擇--");
        while (get_DDL.Read()) {
            ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
            DDL.Items.Add(li);
        }
        get_DDL.Close();
        conn.Close();
        for (int i = 0; i < DDL.Items.Count; i++) {
            if (DDL.Items[i].Text == type1) {
                DDL.SelectedIndex = i;
                DDL.Items[i].Selected = true;
            }
        }
    } 

    protected void SetDDL2(DropDownList DDL, string type2)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        string strSQL = "Select * from type where 1 = 1 AND datalevel = 2 and dataparent='" + DDL_Class1.SelectedValue.ToString() + "'";
        conn.Open();
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        SqlDataReader get_DDL = cmd.ExecuteReader();
        while (get_DDL.Read()) {
            ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
            DDL_Class2.Items.Add(li);

        }
        get_DDL.Close();
        conn.Close();
        for (int i = 0; i < DDL.Items.Count; i++)
        {
            if (DDL.Items[i].Text == type2)
            {
                DDL.SelectedIndex = i;
                DDL.Items[i].Selected = true;
            }

        }
    }

    //第二階下拉
    protected void DDL_Class1_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (DDL_Class1.SelectedItem.Text == "--請選擇--") {
            DDL_Class2.Items.Clear();
            DDL_Class2.Visible = false;
        }
        else {
            DDL_Class2.Visible = true;
            DDL_Class2.Items.Clear();
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            string strSQL = "Select * from type where 1=1 " + "AND datalevel = 2 and dataparent='" + DDL_Class1.SelectedValue.ToString() + "'";
            conn.Open();
            SqlCommand cmd = new SqlCommand(strSQL, conn);
            SqlDataReader get_DDL = cmd.ExecuteReader();
            while (get_DDL.Read()) {
                ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
                DDL_Class2.Items.Add(li);

            }
            get_DDL.Close();
            conn.Close();            
        }
    } 

    protected void insert_first(string userid)
    {
        string yesorno = "N";
        if (Rb_yes.Checked == true) {
            yesorno = "Y";
        }
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        conn.Open();

        string strSQL = "UPDATE CatTreeRoot SET CtRootName = @Rootname, Purpose = @Purpose, inUse = @yesorno WHERE CtRootID = @id_no";

        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Rootname", SqlDbType.NVarChar).Value = Txt_Name.Text;
        cmd.Parameters.Add("@Purpose", SqlDbType.NVarChar).Value = Txt_Name.Text;
        cmd.Parameters.Add("@yesorno", SqlDbType.NChar).Value = yesorno;
        cmd.Parameters.Add("@id_no", SqlDbType.Int).Value = Convert.ToInt32(id_no);
        cmd.ExecuteNonQuery();
        conn.Close();
    }

    protected void insert_second(string user_id)
    {
        Random x = new Random();
        for (int i = 0; i < 10; i++) {
            string num = x.Next(10).ToString();
            nums = nums + num;
        }
        if (Fileup.HasFile) {
            string path = Server.MapPath(dbconfig.Filepath());
            string fileN = Fileup.FileName;
            string subfilename = System.IO.Path.GetExtension(fileN);
            string file_name = "title" + nums + subfilename;

            Fileup.SaveAs(path + file_name);
            SqlConnection conn1 = new SqlConnection(dbconfig.ConnectionSettings());
            conn1.Open();
            string picdata = "SELECT pic FROM NodeInfo WHERE CtRootID = @Rootid";
            SqlCommand cmd1 = new SqlCommand(picdata, conn1);
            cmd1.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
            SqlDataReader reader = cmd1.ExecuteReader();
            while (reader.Read()) {
                string delpath = Server.MapPath(dbconfig.Filepath()) + reader["pic"].ToString();
                System.IO.File.Delete(delpath);
            }
            reader.Close();
            string modpic = "UPDATE NodeInfo SET pic = @Filename WHERE CtRootID = @Rootid";
            SqlCommand modcmd = new SqlCommand(modpic, conn1);
            modcmd.Parameters.Add("@Filename", SqlDbType.VarChar).Value = file_name;
            modcmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
            modcmd.ExecuteNonQuery();
            conn1.Close();
        }

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
      
        conn.Open();

        string strSQL = "UPDATE NodeInfo SET sub_title = @Subtitle, url_link = @URL, abstract = @Description,type1 = @type1, " + 
                        "type2 = @type2 WHERE CtRootID = @Rootid";
      
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Subtitle", SqlDbType.VarChar).Value = Txt_SubTitle.Text;
        cmd.Parameters.Add("@URL", SqlDbType.VarChar).Value = Txt_URL.Text;
        cmd.Parameters.Add("@Description", SqlDbType.VarChar).Value = Txt_Description.Value;
        cmd.Parameters.Add("@type1", SqlDbType.VarChar).Value = DDL_Class1.SelectedItem.Text;
        cmd.Parameters.Add("@type2", SqlDbType.VarChar).Value = DDL_Class2.SelectedItem.Text;
        cmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
        cmd.ExecuteNonQuery();
        conn.Close();
    }
  
    protected void get_data(string x)
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "SELECT * FROM CatTreeRoot LEFT JOIN NodeInfo ON CatTreeRoot.CtRootID = NodeInfo.CtRootID WHERE CatTreeRoot.CtRootID = @Rootid";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
        SqlDataReader reader = cmd.ExecuteReader();
        while (reader.Read()) 
        {
            Txt_Name.Text = reader["ctrootname"].ToString();
            Txt_SubTitle.Text = reader["sub_title"].ToString();
            Txt_Description.Value = reader["abstract"].ToString();
            Txt_URL.Text = reader["url_link"].ToString();
            DBimage.Visible = true;        		
            string path1 = x + reader["pic"].ToString();
            DBimage.ImageUrl = "";
            DBimage.ImageUrl =path1;
            if (reader["inuse"].ToString() == "Y") {
                Rb_yes.Checked = true;
            }
            else {
                Rb_no.Checked = true;
            }
            SetDDL(DDL_Class1, reader["type1"].ToString());
            SetDDL2(DDL_Class2, reader["type2"].ToString());
        }
        reader.Close();
        conn.Close();
    }

    protected void clear()
    {
        Txt_Name.Text = "";
        Txt_SubTitle.Text = "";
        Txt_Description.Value = "";
        Txt_URL.Text = "";
        DBimage.Visible = false;
        DDL_Class1.Items.Clear();
        DDL_Class2.Items.Clear();
    }
}
