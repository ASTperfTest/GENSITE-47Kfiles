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

public partial class Edit_Step2_Edit : System.Web.UI.Page
{
    public string boxid;
    public string nums;
    public string[] bannername;
    public string outletd;
    public string wstyled;
    public string bannerd;
    public Setting dbconfig = new Setting();

    protected void Page_Load(object sender, EventArgs e)
    {
        boxid = Session["User_id"].ToString();
        string name = Session["Name"].ToString();
        if (Request.QueryString["pic"] != null)
        {
            string picname = Request.QueryString["pic"].ToString();
            string deletepath = Server.MapPath("../"+dbconfig.Filepath()) + picname;
            SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
            conn.Open();
            string strSQL = "delete from nodebanner where pic =@picname";
            SqlCommand cmd = new SqlCommand(strSQL, conn);
            cmd.Parameters.Add("@picname", SqlDbType.VarChar).Value = picname;

            cmd.ExecuteNonQuery();
            System.IO.File.Delete(deletepath);
            conn.Close();
            clear();
            get_data();
        }
        if (!IsPostBack)
        {

            set_style1();
            set_style2();
            set_bannerlist();
            get_data();
        }

    }

    protected void set_style1()
    {
        //  string[] files = Directory.GetFiles(@"./images", "outlet*.*");
        string[] files = Directory.GetFiles(Request.PhysicalApplicationPath + dbconfig.Stylepath(), "outlet*.*");
        int i = 0;
        string filename = "";

        while (i < files.Length)
        {
            FileInfo finfo = new FileInfo(files[i]);
            ListItem li = new ListItem();

            li.Text = "<p><Img src='" + dbconfig.EditStylepath() + finfo.Name.ToString() + "' align='left'></p>";

            filename = finfo.Name.ToString();

            filename = filename.Substring(0, filename.IndexOf("."));

            li.Value = filename;

            //li.Value = finfo.Name.ToString().Substring(0, 7);
            //Response.Write(li.Text);
            //Response.Write(li.Value);
            outlet.Items.Add(li);
            i++;

        }
        outlet.Items[0].Selected = true;

    }

    protected void set_style2()
    {
        string[] files = Directory.GetFiles(Request.PhysicalApplicationPath + dbconfig.Stylepath(), "style*.*");
        int i = 0;
        string filename = "";

        while (i < files.Length)
        {
            FileInfo finfo = new FileInfo(files[i]);
            ListItem li = new ListItem();

            li.Text = "<p><Img src='" + dbconfig.EditStylepath() + finfo.Name.ToString() + "' align='middle'></p>";

            filename = finfo.Name.ToString();

            filename = filename.Substring(0, filename.IndexOf("."));
            
            li.Value = filename;

            //li.Value = finfo.Name.ToString().Substring(0, 6);
            //Response.Write(li.Text);
            //Response.Write(li.Value);
            wstyle.Items.Add(li);
            i++;
        }
        wstyle.Items[0].Selected = true;



    }

    protected void set_bannerlist()
    {

        string[] files = Directory.GetFiles(Request.PhysicalApplicationPath + dbconfig.Stylepath(), "banner-*.*");

        int i = 0;
        for (int j = 0; j < files.Length; j++)
        {
            FileInfo finfo = new FileInfo(files[j]);
            string[] splitstring = finfo.ToString().Split('-');
            ListItem li = new ListItem();
            li.Text = "<p>" + splitstring[1].ToString() + "<br/><Img src='" + dbconfig.EditStylepath() + finfo.Name.ToString() + "' align='middle' width=250 height=60></p>";
            li.Value = finfo.Name.ToString();
            bannerlist.Items.Add(li);
            i++;
        }
        bannerlist.Items[0].Selected = true;


        string path = Session["PublicPath"].ToString();
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        conn.Open();
        string getbanner = "select * from nodebanner where ctrootid = '" + boxid + "'";
        SqlCommand getbanner1 = new SqlCommand(getbanner, conn);
        getbanner1.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
        SqlDataReader reader = getbanner1.ExecuteReader();
        
        while (reader.Read())
        {
            ListItem li = new ListItem();
            li.Text = "<p>" + reader["title"].ToString() +"  /  "+ "<a href=step2_edit.aspx?pic="+reader["pic"].ToString()+">刪除</a><br/><Img src='" + path + reader["pic"].ToString() + "' align='middle' width=250 height=60></p>";
            li.Value = reader["pic"].ToString();
            bannerlist.Items.Add(li);
            
            i++;
            bannerlist.Items[i-1].Selected = true;
        }
       

    }

    protected void next_Click(object sender, EventArgs e)
    {
        Upload_topic();

        foreach (ListItem item in outlet.Items)
        {
            if (item.Selected)
            {
                outletd = item.Value;

            }
        }

        foreach (ListItem item in wstyle.Items)
        {
            if (item.Selected)
            {
                wstyled = item.Value;
            }
        }

        foreach (ListItem item in bannerlist.Items)
        {
            if (item.Selected)
            {
                bannerd = item.Value;

            }
        }

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        conn.Open();
        string writedata1 = "update nodeinfo set list =@outletd ,data=@wstyled ,pic_banner=@bannerd where ctrootid =@Rootid";
        // string writedata2 = "update nodeinfo set pic_banner = '" + bannerd + "' where ctnodeid ='" + boxid + "'";
        SqlCommand updatedata = new SqlCommand(writedata1, conn);
        updatedata.Parameters.Add("@outletd", SqlDbType.VarChar).Value = outletd;
        updatedata.Parameters.Add("@wstyled", SqlDbType.VarChar).Value = wstyled;
        updatedata.Parameters.Add("@bannerd", SqlDbType.VarChar).Value = bannerd;
        updatedata.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
        //  SqlCommand updatedata2 = new SqlCommand(writedata2, conn);


        updatedata.ExecuteNonQuery();
        // updatedata2.ExecuteNonQuery();
        Session.Add("style", wstyled);

        conn.Close();
        Response.Redirect("../GIP/web/step3.aspx");





    }

    protected void Upload_topic()
    {
        Random x = new Random();
        for (int i = 0; i < 10; i++)
        {
            string num = x.Next(10).ToString();
            nums = nums + num;
        }
        if (Topic_upload.HasFile)
        {
            string path = Server.MapPath("../" + dbconfig.EditFilepath()) ;
            string fileN = Topic_upload.FileName;
            string subfilename = System.IO.Path.GetExtension(fileN);
            string file_name = "Topic" + nums + subfilename;

            Topic_upload.SaveAs(path + file_name);

            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            conn.Open();
            string dataf = "select * from nodeinfo where ctrootid =@Rootid";
            SqlCommand dataff = new SqlCommand(dataf, conn);
            dataff.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
            SqlDataReader readcc = dataff.ExecuteReader();
            if (readcc.Read())
            {
                if (!Convert.IsDBNull(readcc["pic_title"]))
                {
                    string oldfile = readcc["pic_title"].ToString();
                    string deletepath = Server.MapPath("../"+ dbconfig.EditFilepath()) + oldfile;
                    System.IO.File.Delete(deletepath);
                    readcc.Close();
                    string writedata3 = "update nodeinfo set pic_title =@Filename where ctrootid =@Rootid";
                    SqlCommand updatepic = new SqlCommand(writedata3, conn);
                    updatepic.Parameters.Add("@Filename", SqlDbType.VarChar).Value = file_name;
                    updatepic.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
                    updatepic.ExecuteNonQuery();
                }
                else
                {
                    readcc.Close();
                    string writedata3 = "update nodeinfo set pic_title =@Filename where ctrootid =@Rootid";
                    SqlCommand updatepic = new SqlCommand(writedata3, conn);
                    updatepic.Parameters.Add("@Filename", SqlDbType.VarChar).Value = file_name;
                    updatepic.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
                    updatepic.ExecuteNonQuery();

                }
                readcc.Close();

            }
            conn.Close();
        }
    }

    protected void save_Click(object sender, EventArgs e)
    {
        Upload_topic();

        foreach (ListItem item in outlet.Items)
        {
            if (item.Selected)
            {
                outletd = item.Value;

            }
        }

        foreach (ListItem item in wstyle.Items)
        {
            if (item.Selected)
            {
                wstyled = item.Value;
            }
        }

        foreach (ListItem item in bannerlist.Items)
        {
            if (item.Selected)
            {
                bannerd = item.Value;

            }
        }

        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        conn.Open();
        string writedata1 = "update nodeinfo set list =@outletd ,data=@wstyled ,pic_banner =@bannerd where ctrootid =@Rootid";
        //string writedata2 = "update nodeinfo set pic_banner = '" + bannerd + "' where ctnodeid ='" + boxid + "'";
        SqlCommand updatedata = new SqlCommand(writedata1, conn);
        updatedata.Parameters.Add("@outletd", SqlDbType.VarChar).Value = outletd;
        updatedata.Parameters.Add("@wstyled", SqlDbType.VarChar).Value = wstyled;
        updatedata.Parameters.Add("@bannerd", SqlDbType.VarChar).Value = bannerd;
        updatedata.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
        //SqlCommand updatedata2 = new SqlCommand(writedata2, conn);


        updatedata.ExecuteNonQuery();
        // updatedata2.ExecuteNonQuery();

        conn.Close();
        clear();
        get_data();



    }
    protected void delete_Click(object sender, EventArgs e)
    {
        SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
        conn.Open();
        string get_data = "select * from nodebanner where ctrootid=@Rootid";
        SqlCommand deletecmd1 = new SqlCommand(get_data, conn);
        deletecmd1.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
        SqlDataReader reader = deletecmd1.ExecuteReader();
        while (reader.Read())
        {
            string filename = reader["pic"].ToString();
            string deletepath = Server.MapPath("../" + dbconfig.EditFilepath()) + filename;
            System.IO.File.Delete(deletepath);


        }
        reader.Close();

        string delete_data = "delete from nodebanner where ctrootid=@Rootid";
        SqlCommand deletecmd2 = new SqlCommand(delete_data, conn);
        deletecmd2.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
        deletecmd2.ExecuteNonQuery();

        string get_data3 = "select * from nodeinfo where ctrootid=@Rootid";
        SqlCommand deletecmd3 = new SqlCommand(get_data3, conn);
        deletecmd3.Parameters.Add("@Rootid", SqlDbType.Int).Value = Convert.ToInt32(boxid);
        SqlDataReader reader3 = deletecmd3.ExecuteReader();
        while (reader3.Read())
        {
            if (!Convert.IsDBNull(reader3["pic"]))
            {
                string titlename = reader3["pic"].ToString();
                string deletetitle = Server.MapPath("../" + dbconfig.EditFilepath()) + titlename;
                System.IO.File.Delete(deletetitle);
            }

            if (!Convert.IsDBNull(reader3["pic_title"]))
            {
                string topicname = reader3["pic_title"].ToString();
                string deletetopic = Server.MapPath("../" + dbconfig.EditFilepath()) + topicname;
                System.IO.File.Delete(deletetopic);
            }
        }
        reader3.Close();

        string delete_data2 = "delete from nodeinfo where ctrootid=" + boxid ;
        SqlCommand deletedatatable = new SqlCommand(delete_data2, conn);
        string delete_data3 = "delete from cattreeroot where ctrootid =" + boxid;
        SqlCommand deletedatatable2 = new SqlCommand(delete_data3, conn);
        
        deletedatatable.ExecuteNonQuery();
        deletedatatable2.ExecuteNonQuery();
        conn.Close();
     
       Response.Redirect("../index.aspx");

    }
    protected void back_Click(object sender, EventArgs e)
    {
        Response.Redirect("./Step1_Edit.aspx");
    }

    protected void get_data()
    {
        SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
        conn.Open();
        string strSQL = "select * from nodeinfo where ctrootid ='"+boxid+"'";
        SqlCommand cmd = new SqlCommand(strSQL, conn);
        SqlDataReader reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            foreach (ListItem item in outlet.Items)
            {
                if (reader["list"].ToString() == item.Value)
                {
                    item.Selected = true;

                }
            }

            foreach (ListItem item in wstyle.Items)
            {
                if (reader["data"].ToString() == item.Value)
                {
                    //wstyled = item.Value;
                    item.Selected = true;
                }
            }

            foreach (ListItem item in bannerlist.Items)
            {
                if (reader["pic_banner"].ToString() == item.Value)
                {
                    //bannerd = item.Value;
                    item.Selected = true;

                }
            }
            if(!Convert.IsDBNull(reader["pic_title"]))
            {
                Topic_pic.ImageUrl = Session["PublicPath"].ToString() + reader["pic_title"].ToString();
            }
        }

    
    }
    protected void clear()
    {
        outlet.ClearSelection();
        wstyle.ClearSelection();
        bannerlist.ClearSelection();

        
    
    }
    protected void cancel_Click(object sender, EventArgs e)
    {
        Response.Redirect("../index.aspx");
    }
}
