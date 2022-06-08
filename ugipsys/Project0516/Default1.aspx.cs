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
       // Class1 xxx = new Class1();
        //public Setting = new Setting();
        public Setting dbconfig = new Setting();

        protected void Page_Load(object sender, EventArgs e)
        {
           
            user_idno = Session["Name"].ToString();
            //Response.Write(Session["Name"].ToString());
            //Response.Write(user_idno);
          //  string path = System.Web.HttpContext.Current.Server.MapPath("../public/finish.aspx");
           // Response.Write(path);
            if (!IsPostBack)
            {
                SetDDL(DDL_Class1);
            }
        }

        protected void next_Click(object sender, EventArgs e)
        {


            Response.Redirect("step1_2.aspx");
        } //下一步

        protected void SetDDL(DropDownList DDL)
        {
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            string strSQL = "Select * from type where 1=1" + "AND datalevel = 1";
            conn.Open();
            SqlCommand cmd = new SqlCommand(strSQL, conn);
            SqlDataReader get_DDL = cmd.ExecuteReader();
            DDL.Items.Add("--請選擇--");
            while (get_DDL.Read())
            {
                ListItem li = new ListItem(get_DDL["classname"].ToString(), get_DDL["classid"].ToString());
                DDL.Items.Add(li);
            }
            get_DDL.Close();
            conn.Close();
        } //設定下拉

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
                //Response.Write(DDL_Class1.SelectedValue);
                //Response.Write(DDL_Class1.SelectedItem.Value);
            }
        } //第二階下拉

        protected void insert_first(string userid)
        {
            //string catname = Session["Name"].ToString() + Txt_Name.Text;
            string yesorno;
            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            conn.Open();
            //string getid = "select * from cattreenode where catname ='" + Session["Name"].ToString() + "'";
            //SqlCommand get_id = new SqlCommand(getid, conn);
            //SqlDataReader reader = get_id.ExecuteReader();
            //while (reader.Read())
            //{
            //    userid = reader["ctNodeId"].ToString();
            //}
            //id_no = userid;
            //reader.Close();
            if (Rb_yes.Checked == true)
            {
                yesorno = "Y";
            }
            else
            {
                yesorno = "N";
            }

            string strSQL = "insert into cattreeroot(ctrootname,purpose,inuse,editor,vgroup) values('" + Txt_Name.Text + "'," + "'" + Txt_Name.Text + "','" + yesorno + "','" + user_idno + "','G1')";
                SqlCommand cmd = new SqlCommand(strSQL, conn);
                cmd.ExecuteNonQuery();

                string getid = "select * from cattreeroot where 1=1" + "and ctrootname ='" + Txt_Name.Text + "'" + " and editor='" + user_idno + "'";
                SqlCommand cmd1 = new SqlCommand(getid, conn);
                SqlDataReader zzz = cmd1.ExecuteReader();
                //Response.Write(getid);
                while (zzz.Read())
                {
                    id_no = zzz["ctRootid"].ToString();
                    Session.Add("User_id", id_no);
                    //Response.Write(Session["User_id"].ToString());
                }
                zzz.Close();

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

            string path = Server.MapPath(dbconfig.Filepath());
            string fileN = Fileup.FileName;
            string subfilename = System.IO.Path.GetExtension(fileN);
            string file_name = "title" + nums + subfilename;

            Fileup.SaveAs(path + file_name);

            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);

            // string getkey = "select * from cattreenode where dataparent = '" + user_id + "' and catname ='" + Txt_Name.Text + "'";

             conn.Open();

            // SqlCommand gkey = new SqlCommand(getkey, conn);
            // SqlDataReader reader = gkey.ExecuteReader();

            // while (reader.Read())
            // {
            //     user_id = reader["ctNodeId"].ToString();
            //     Session.Add("Boxid", user_id);

            // }
            // reader.Close();
           

                string strSQL = "insert into nodeinfo(ctrootid,sub_title,url_link,abstract,type1,type2,pic,owner) values(@Rootid,@Subtitle,@URL,"+
                    "@Description,@type1,@type2,@pic,@owner)";
                
                //id_no + "','" + Txt_SubTitle.Text + "','" + Txt_URL.Text + "','" + Txt_Description.Value + "','" + DDL_Class1.SelectedItem.Text + "','" +
                //DDL_Class2.SelectedItem.Text + "','" + file_name + "','"+ user_idno +"')";

                SqlCommand cmd = new SqlCommand(strSQL, conn);
                //da.SelectCommand.Parameters.Add("@TreeRootId", SqlDbType.Int).Value = TreeRootId;   
                cmd.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(id_no.ToString());
                cmd.Parameters.Add("@Subtitle", SqlDbType.VarChar).Value = Txt_SubTitle.Text;
                cmd.Parameters.Add("@URL", SqlDbType.VarChar).Value = Txt_URL.Text;
                cmd.Parameters.Add("@Description", SqlDbType.VarChar).Value = Txt_Description.Value;
                cmd.Parameters.Add("@type1", SqlDbType.VarChar).Value = DDL_Class1.SelectedItem.Text;
                cmd.Parameters.Add("@type2", SqlDbType.VarChar).Value = DDL_Class2.SelectedItem.Text;
                cmd.Parameters.Add("@pic", SqlDbType.VarChar).Value = file_name;
                cmd.Parameters.Add("@owner", SqlDbType.VarChar).Value = user_idno;
                cmd.ExecuteNonQuery();
                conn.Close();
                Session.Add("type1", DDL_Class1.SelectedItem.Value.ToString());
           
        }

        
        protected void save_Click(object sender, EventArgs e)
        {
            if (Txt_URL.Text == "")
            {
                insert_first(user_idno);
                insert_second(user_idno);
                next.Visible = true;
                save.Visible = false;
                delete.Visible = true;
                Response.Redirect("./step1_2.aspx");
                
            }
            else
            {
                insert_first(user_idno);
                insert_second(user_idno);
                next.Visible = false;
                save.Visible = false;
                //Response.Write("<script language=\"javascript\">window.onload=function(){alert(\"新增成功!\");window.close();opener.location.reload();}</script>");
                Response.Redirect("./Finish.aspx");
            }
        }
        protected void delete_Click(object sender, EventArgs e)
        {
           // string delSQL = "delete from cattreeroot where ctrootid ='"+id_no+"'";
          //  string delSQL2 = "delete from nodeinfo where ctrootid='" + id_no + "'";
             id_no = Session["User_id"].ToString();
             string delSQL = "delete from cattreeroot where ctrootid ="+id_no;
             string delSQL2 = "delete from nodeinfo where ctrootid="+id_no;

             dbconfig.deletetitle(delSQL, delSQL2, id_no, Server.MapPath(dbconfig.Filepath()),Session["GipDsdPath"].ToString());
          
        }
        protected void back_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/index.aspx");
        }
}
