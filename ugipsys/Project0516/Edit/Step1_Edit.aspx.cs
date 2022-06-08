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
using System.IO;

public partial class Edit_Step1_Edit : System.Web.UI.Page
{
    public string id_no;
    public string nums;
    public string user_idno;
    public string choice;
	public string kmKeyword;
    // Class1 xxx = new Class1();
    //public Setting = new Setting();
    public Setting dbconfig = new Setting();
	public Boolean isAdmin = false;

    protected void Page_Load(object sender, EventArgs e)
    {	
	    string x = Session["PublicPath"].ToString();
      if (Request.QueryString["ID"] != null)
      {
        id_no = Request.QueryString["ID"].ToString();
        Session["User_id"] = id_no;
      }
      else
      {
        id_no = Session["User_id"].ToString();
      }
      user_idno = Session["Name"].ToString();
	
      if (!IsPostBack)
      {
        get_data(x);
        Session.Add("Genname", Txt_Name.Text);
      }
	  //判斷是不是admin user
	  SqlConnection conn = null;
	  SqlDataReader reader = null;
	  string userRight = "SELECT ugrpID FROM InfoUser WHERE UserID = '" + Session["Name"].ToString() + "'";
      conn = new SqlConnection(dbconfig.ConnectionSettings());
      conn.Open();
      SqlCommand comm = new SqlCommand(userRight, conn);
      reader = comm.ExecuteReader();
      if ( reader.Read() ) {
        if (reader["ugrpID"].ToString().Contains("HTSD") || reader["ugrpID"].ToString().Contains("SysAdm") ) {
          isAdmin = true;
        }
      }
      reader.Close();
      Txt_order.Visible = isAdmin;
	  //判斷是不是admin user  end
	  kmKeyword = GetCoaKeywords();
	  
    }

    protected void next_Click(object sender, EventArgs e)
    {
        if (Txt_URL.Text == "")
        {
            insert_first(user_idno);
            insert_second(user_idno);
            next.Visible = true;
            //save.Visible = false;
            delete.Visible = true;
            if (Session["Genname"] == null)
            {
                Session.Add("Genname", Txt_Name.Text);
            }
            else
            {
                Session["Genname"] = Txt_Name.Text;
            }
            Response.Redirect("./step2_edit.aspx");
            

        }
        else
        {
            insert_first(user_idno);
            insert_second(user_idno);
            //next.Visible = false;
            //save.Visible = false;
            //Response.Write("<script language=\"javascript\">window.onload=function(){alert(\"\");window.close();opener.location.reload();}</script>");
            Session.Add("URL", "y");
            Response.Redirect("../Finish.aspx");
        }
    } //

    protected void SetDDL(DropDownList DDL,string type1)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        string strSQL = "Select * from type where 1=1" + "AND datalevel = 1";
        conn.Open();
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        SqlDataReader get_DDL = cmd.ExecuteReader();
        //DDL.Items.Add("--請選擇--");
        while (get_DDL.Read())
        {
            ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
            DDL.Items.Add(li);
            
        }
        get_DDL.Close();
        conn.Close();
        for (int i = 0; i < DDL.Items.Count; i++)
        {
            if (DDL.Items[i].Text == type1)
            {
                DDL.SelectedIndex = i;
                DDL.Items[i].Selected = true;
            }
        
        }

    } //

    protected void SetDDL2(DropDownList DDL, string type2)
    {
        
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        string strSQL = "Select * from type where 1=1 " + "AND datalevel = 2 and dataparent='" + DDL_Class1.SelectedValue.ToString() + "'";
        conn.Open();
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        SqlDataReader get_DDL = cmd.ExecuteReader();
        while (get_DDL.Read())
        {
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

    protected void DDL_Class1_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (DDL_Class1.SelectedItem.Text == "--請選擇--")
        {
            DDL_Class2.Items.Clear();
            DDL_Class2.Visible = false;
        }
        else
        {
            DDL_Class2.Visible = true;
            DDL_Class2.Items.Clear();
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            string strSQL = "Select * from type where 1=1 " + "AND datalevel = 2 and dataparent=@dataparent";
            conn.Open();
            SqlCommand cmd = new SqlCommand(strSQL, conn);
            cmd.Parameters.Add("@dataparent", SqlDbType.Int).Value = DDL_Class1.SelectedValue;
            SqlDataReader get_DDL = cmd.ExecuteReader();
            while (get_DDL.Read())
            {
                ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
                DDL_Class2.Items.Add(li);
                
            }
            get_DDL.Close();
            conn.Close();
            //Response.Write(DDL_Class1.SelectedValue);
            //Response.Write(DDL_Class1.SelectedItem.Value);
        }
    } //

    protected void insert_first(string userid)
    {
        
        string yesorno;
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        conn.Open();
        
        if (Rb_yes.Checked == true)
        {
            yesorno = "Y";
        }
        else
        {
            yesorno = "N";
        }


        string strSQL = "update cattreeroot set ctrootname =@Rootname, purpose=@Purpose, inuse=@yesorno where ctrootid =@id_no";

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
        for (int i = 0; i < 10; i++)
        {
            string num = x.Next(10).ToString();
            nums = nums + num;
        }
        if (Fileup.HasFile)
        {
            string path = Server.MapPath("../" + dbconfig.EditFilepath());
            string fileN = Fileup.FileName;
            string subfilename = System.IO.Path.GetExtension(fileN);
            string file_name = "title" + nums + subfilename;

            Fileup.SaveAs(path + file_name);
            SqlConnection conn1 = new SqlConnection(dbconfig.ConnectionSettings());
            conn1.Open();
            string picdata = "select pic from nodeinfo where ctrootid=@Rootid";
            SqlCommand cmd1 = new SqlCommand(picdata, conn1);
            cmd1.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
            SqlDataReader reader = cmd1.ExecuteReader();
            while (reader.Read())
            {
                string delpath =Server.MapPath("../" + dbconfig.EditFilepath()) + reader["pic"].ToString();
                System.IO.File.Delete(delpath);
            }
            reader.Close();
            string modpic = "update nodeinfo set pic = @Filename where ctrootid=@Rootid";
            SqlCommand modcmd = new SqlCommand(modpic, conn1);
            modcmd.Parameters.Add("@Filename", SqlDbType.VarChar).Value = file_name;
            modcmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
            modcmd.ExecuteNonQuery();
            conn1.Close();
        }
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);

        // string getkey = "select * from cattreenode where dataparent = '" + user_id + "' and catname ='" + Txt_Name.Text + "'";

        conn.Open();        
        string strSQL = "update nodeinfo set sub_title =@Subtitle,url_link=@URL,abstract=@Description,type1=" +
            "@type1,type2=@type2, lockRightBtn = @lockrightbtn, footer_dept = @footer_dept, footer_dept_url =" +
			"@footer_dept_url, order_num = @order_num,Keywords=@Keywords,CatMemo_Disable = @CatMemo_Dis  where ctrootid =@Rootid";
        // Response.Write(strSQL);
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Subtitle", SqlDbType.VarChar).Value = Txt_SubTitle.Text;
        cmd.Parameters.Add("@URL", SqlDbType.VarChar).Value = Txt_URL.Text;
        cmd.Parameters.Add("@Description", SqlDbType.VarChar).Value = Txt_Description.Value;
        cmd.Parameters.Add("@type1", SqlDbType.VarChar).Value = DDL_Class1.SelectedItem.Text;
        cmd.Parameters.Add("@type2", SqlDbType.VarChar).Value = DDL_Class2.SelectedItem.Text;
        cmd.Parameters.Add("@lockrightbtn", SqlDbType.Char).Value = lockRightBtnDDL.SelectedValue;
        cmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
		cmd.Parameters.Add("@footer_dept", SqlDbType.VarChar).Value = Txt_dept.Value;
		cmd.Parameters.Add("@footer_dept_url", SqlDbType.VarChar).Value = Txt_dept_URL.Text;
		cmd.Parameters.Add("@order_num", SqlDbType.VarChar).Value = Txt_order.Text;
		cmd.Parameters.Add("@Keywords", SqlDbType.VarChar).Value = (Txt_Keywords.Text).Replace(" ","");
		cmd.Parameters.Add("@CatMemo_Dis", SqlDbType.VarChar).Value = ((CheckBox)FindControl("CatMemo_Disable")).Checked;
        cmd.ExecuteNonQuery();
        conn.Close();

    }


    protected void save_Click(object sender, EventArgs e)
    {
        if (Txt_URL.Text == "")
        {
            insert_first(user_idno);
            insert_second(user_idno);
            clear();
            get_data(Session["PublicPath"].ToString());
            
        }
        else
        {
            insert_first(user_idno);
            insert_second(user_idno);
            clear();
            get_data(Session["PublicPath"].ToString());
           // next.Visible = false;
            //save.Visible = false;
            //Response.Write("<script language=\"javascript\">window.onload=function(){alert(\"\");window.close();opener.location.reload();}</script>");
            
        }
    }

    //---刪除主題館---
    protected void delete_Click(object sender, EventArgs e)
    {
        //string delSQL = "delete from cattreeroot where ctrootid ='"+id_no+"'";
        //string delSQL2 = "delete from nodeinfo where ctrootid='" + id_no + "'";
        //id_no = Session["User_id"].ToString();
	      //id_no = "42";

	      //string test = Session["User_id"].ToString();
        string id = Request.QueryString["id"];
        
        string delSQL = "delete from cattreeroot where ctrootid ='" + id + "'";
        string delSQL2 = "delete from nodeinfo where ctrootid='" + id + "'";
        dbconfig.deletetitle(delSQL, delSQL2, id, Server.MapPath("../" + dbconfig.EditFilepath()), Session["GipDsdPath"].ToString());
	      Response.Redirect("../index.aspx");
    }

    protected void get_data(string x)
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "select * from cattreeroot left join nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where cattreeroot.ctrootid =@Rootid";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        cmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
        SqlDataReader reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            Txt_Name.Text = reader["ctrootname"].ToString();
            Txt_SubTitle.Text = reader["sub_title"].ToString();
            Txt_Description.Value = reader["abstract"].ToString();
	        Txt_URL.Text = reader["url_link"].ToString();
			Txt_dept.Value = reader["footer_dept"].ToString();
			Txt_dept_URL.Text = reader["footer_dept_url"].ToString();
			Txt_order.Text = reader["order_num"].ToString();
			Txt_Keywords.Text = reader["Keywords"].ToString();
			if(reader["lockrightbtn"].ToString()!= "" && reader["lockrightbtn"].ToString() =="N" )
			{
				lockRightBtnDDL.SelectedIndex = 1;
			}
			if(reader["CatMemo_Disable"].ToString() != "" && reader["CatMemo_Disable"] != null)
			{
				((CheckBox)FindControl("CatMemo_Disable")).Checked = (bool)reader["CatMemo_Disable"];
			}
            DBimage.Visible = true;
           // string path = Session["PublicPath"].ToString() + reader["pic"].ToString();
string path = x + reader["pic"].ToString();
            DBimage.ImageUrl = path;
            if (reader["inuse"].ToString() == "Y")
            {
                Rb_yes.Checked = true;
            }
            else
            {
                Rb_no.Checked = true;
            }

            SetDDL(DDL_Class1,reader["type1"].ToString());
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
    protected void back_Click(object sender, EventArgs e)
    {
        Response.Redirect("../index.aspx");
    }
	
	private string GetCoaKeywords()
	{
	    string keywordTemp="";
	    SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["KM_ConnectionString"].ConnectionString);
		conn.Open();
		string strSQL = "SELECT TOP 10 KEYWORD FROM REPORT_KEYWORD_FREQUENCY  WHERE MP = "+id_no+"GROUP BY KEYWORD ORDER BY SUM(FREQUENCY) DESC ";
		SqlCommand cmd = new SqlCommand(strSQL, conn);
		SqlDataReader reader = cmd.ExecuteReader();
		    while(reader.Read())
			{
			     if(keywordTemp != "") keywordTemp +=",";
			     keywordTemp+=reader["KEYWORD"].ToString();
			}
		conn.Close();
		return keywordTemp;
	}
}
