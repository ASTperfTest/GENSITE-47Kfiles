using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;

    public partial class _Default : System.Web.UI.Page
    {
        
        public string id_no;
        public string nums;
        public string user_idno;
        public Setting dbconfig = new Setting();

        protected void Page_Load(object sender, EventArgs e)
        {           
            user_idno = Session["Name"].ToString();
            if (!IsPostBack) {
                SetDDL(DDL_Class1);
            }
        }
        //下一步
        protected void next_Click(object sender, EventArgs e)
        {
            Response.Redirect("step1_2.aspx");
        }
        //暫存主題館
        protected void save_Click(object sender, EventArgs e)
        {
          if (Txt_URL.Text == "") {
            insert_first(user_idno);
            insert_second(user_idno);
            next.Visible = true;
            save.Visible = false;
            delete.Visible = true;
            Response.Redirect("./step1_2.aspx");
          }
          else {
            insert_first(user_idno);
            insert_second(user_idno);
            next.Visible = false;
            save.Visible = false;
            //Response.Write("<script language=\"javascript\">window.onload=function(){alert(\"新增成功!\");window.close();opener.location.reload();}</script>");
            Session.Add("URL", "y");
            Response.Redirect("./Finish.aspx");
          }
        }
        //刪除主題館
        protected void delete_Click(object sender, EventArgs e)
        {
          // string delSQL = "delete from cattreeroot where ctrootid ='"+id_no+"'";
          //  string delSQL2 = "delete from nodeinfo where ctrootid='" + id_no + "'";
          id_no = Session["User_id"].ToString();
          string delSQL = "delete from cattreeroot where ctrootid =" + id_no;
          string delSQL2 = "delete from nodeinfo where ctrootid=" + id_no;

          dbconfig.deletetitle(delSQL, delSQL2, id_no, Server.MapPath(dbconfig.Filepath()), Session["GipDsdPath"].ToString());

        }
        //回首頁
        protected void back_Click(object sender, EventArgs e)
        {
          Response.Redirect("~/index.aspx");
        }
        //設定下拉
        protected void SetDDL(DropDownList DDL)
        {
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            SqlCommand cmd = null;
            SqlDataReader get_DDL = null;
            string strSQL = "";
            try {
              strSQL = "Select * from type where 1=1" + "AND datalevel = 1";
              conn.Open();
              cmd = new SqlCommand(strSQL, conn);
              get_DDL = cmd.ExecuteReader();              
              while (get_DDL.Read()) {
                ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
                DDL.Items.Add(li);
              }
              get_DDL.Close();
            }
            catch (Exception ex) { }
            finally {
              conn.Close();
            }
            conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            try {
              DDL_Class2.Visible = true;
              DDL_Class2.Items.Clear();
              strSQL = "Select * from type where 1=1 " + "AND datalevel = 2 and dataparent='" + DDL_Class1.SelectedValue.ToString() + "'";
              conn.Open();
              cmd = new SqlCommand(strSQL, conn);
              get_DDL = cmd.ExecuteReader();
              while (get_DDL.Read()) {
                ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
                DDL_Class2.Items.Add(li);
              }
              get_DDL.Close();
            }          
            catch (Exception ex) { }
            finally {
              conn.Close();
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

        // 新增至CatTreeRoot
        protected void insert_first(string userid)
        {           
            string yesorno = "N";
            if (Rb_yes.Checked == true) {
              yesorno = "Y";
            }            
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            try {
              // 新增至CatTreeRoot
              string strSQL = "INSERT INTO CatTreeRoot(CtRootName, Purpose, inUse, Editor,vGroup) VALUES(@Rootname, @Purpose, @Yesorno, @Editor,'XX')";
              conn.Open();
              SqlCommand cmd = new SqlCommand(strSQL, conn);
              cmd.Parameters.Add("@Rootname", SqlDbType.NVarChar).Value = Txt_Name.Text;
              cmd.Parameters.Add("@Purpose", SqlDbType.NVarChar).Value = Txt_Name.Text;
              cmd.Parameters.Add("@Editor", SqlDbType.NVarChar).Value = user_idno;
              cmd.Parameters.Add("@Yesorno", SqlDbType.NChar).Value = yesorno;
              cmd.ExecuteNonQuery();

              string getid = "SELECT * FROM CatTreeRoot WHERE 1 = 1 " + " AND CtRootName = @Rootname AND editor = @Editor";
              SqlCommand cmd1 = new SqlCommand(getid, conn);
              cmd1.Parameters.Add("@Rootname", SqlDbType.NVarChar).Value = Txt_Name.Text;
              cmd1.Parameters.Add("@Editor", SqlDbType.NVarChar).Value = user_idno;
              SqlDataReader zzz = cmd1.ExecuteReader();              
              while (zzz.Read()) {
                id_no = zzz["ctRootid"].ToString();
                Session.Add("User_id", id_no);                
              }
              zzz.Close();

              //---把CtRootId取出, 放到pvXdmp中---若沒有外部連結---
              if (Txt_URL.Text == "")
              {
                string xdmp = "UPDATE CatTreeRoot SET pvxdmp = @Pvxdmp WHERE CtRootID = @Rootid";
                SqlCommand setxdmp = new SqlCommand(xdmp, conn);
                setxdmp.Parameters.Add("@Pvxdmp", SqlDbType.NVarChar).Value = id_no;
                setxdmp.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no);
                setxdmp.ExecuteNonQuery();
              }
            }
            catch (Exception ex) {}
            finally {
              conn.Close();
            }                                    
        }

        protected void insert_second(string user_id)
        {
            Random x = new Random();
            for (int i = 0; i < 10; i++)  {
                string num = x.Next(10).ToString();
                nums = nums + num;
            }

            string path = Server.MapPath(dbconfig.Filepath());
            string fileN = Fileup.FileName;
            string subfilename = System.IO.Path.GetExtension(fileN);
            string file_name = "title" + nums + subfilename;

            Fileup.SaveAs(path + file_name);

            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);

            conn.Open();

            string strSQL = "INSERT INTO NodeInfo(CtrootID, sub_title, url_link, abstract, type1, type2, pic, owner, footer_dept, footer_dept_url, order_num) " +
                            "VALUES( @Rootid, @Subtitle, @URL, @Description, @type1, @type2, @pic, @owner, @footer_dept, @footer_dept_url, @order_num)";
                                
            SqlCommand cmd = new SqlCommand(strSQL, conn);
            cmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no.ToString());
            cmd.Parameters.Add("@Subtitle", SqlDbType.VarChar).Value = Txt_SubTitle.Text;
            cmd.Parameters.Add("@URL", SqlDbType.VarChar).Value = Txt_URL.Text;
            cmd.Parameters.Add("@Description", SqlDbType.VarChar).Value = Txt_Description.Value;
            cmd.Parameters.Add("@type1", SqlDbType.VarChar).Value = DDL_Class1.SelectedItem.Text;
            cmd.Parameters.Add("@type2", SqlDbType.VarChar).Value = DDL_Class2.SelectedItem.Text;
            cmd.Parameters.Add("@pic", SqlDbType.VarChar).Value = file_name;
            cmd.Parameters.Add("@owner", SqlDbType.VarChar).Value = user_idno;
			cmd.Parameters.Add("@footer_dept", SqlDbType.VarChar).Value = Txt_dept.Value;
			cmd.Parameters.Add("@footer_dept_url", SqlDbType.VarChar).Value = Txt_dept_URL.Text;
			cmd.Parameters.Add("@order_num", SqlDbType.VarChar).Value = Txt_order.Text;
            cmd.ExecuteNonQuery();
            conn.Close();
            Session.Add("type1", DDL_Class1.SelectedItem.Value.ToString());           
        }              
    }
