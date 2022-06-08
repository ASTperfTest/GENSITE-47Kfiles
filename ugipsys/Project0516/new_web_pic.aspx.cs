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

public partial class new_web_pic : System.Web.UI.Page
{
    public string id;
    public string nums;
    Setting dbconfig = new Setting();
    protected void Page_Load(object sender, EventArgs e)
    {
        id = Session["User_id"].ToString();
    }

    protected void go_Click(object sender, EventArgs e)
    {

        Random x = new Random();
        for (int i = 0; i < 10; i++)
        {
            string num = x.Next(10).ToString();
            nums = nums + num;
        }

        if (Banner_Upload.HasFile)
        {

            string path = Server.MapPath(dbconfig.Filepath());
            string fileN = Banner_Upload.FileName;
            string subfilename = System.IO.Path.GetExtension(fileN);
            string file_name = "CustomerBanner" + nums + subfilename;

            Banner_Upload.SaveAs(path + file_name);

            SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
            conn.Open();

            
            string writedata3 = "insert into nodebanner(ctrootid,title,pic) values('" + id + "','" + Txt_topic.Text + "','" + file_name + "')";
            SqlCommand updatepic = new SqlCommand(writedata3, conn);
            updatepic.ExecuteNonQuery();
        }
        Response.Write("<script language=\"javascript\">window.onload=function(){alert(\"新增成功!\");window.close();opener.location.reload();}</script>");
    }
}
