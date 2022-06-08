using System;
using System.Collections.Generic;
using System.Data;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.IO;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;

using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Transactions;
using System.Globalization;

public partial class index : Page
{
	private IGardeningService gardeningService;
    
    protected void Page_Load(object sender, EventArgs e)
    {
        //更新目前user的暱稱和realname
		//將入口網的user資訊同步到gardening				
        if( Session["memID"] != null && Session["memID"].ToString().Trim() != string.Empty )
		{		
			UpdateUserInfo(Session["memID"].ToString());
		}
        //LoadHotArticle
        LoadHotArticle();
    }

    private void UpdateUserInfo(string user)
    {
        string nickname = "";
        string realname = "";
		
        //select
        using (SqlConnection tSqlConn = new SqlConnection())
        {
            string webConfigConnectionString1 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

            tSqlConn.ConnectionString = webConfigConnectionString1;
            tSqlConn.Open();

            string selectCommand = "select nickname,realname from member where account=@account ";
            SqlCommand selectCmd = new SqlCommand(selectCommand, tSqlConn);
            selectCmd.Parameters.AddWithValue("account", user);
           
            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommandBuilder cb = new SqlCommandBuilder();
            DataTable dt = new DataTable();

            da.SelectCommand = selectCmd;
            cb.DataAdapter = da; // 指定DataAdapter
            da.Fill(dt);

            if (dt.Rows.Count > 0 && dt != null)
            {
                nickname = dt.Rows[0]["nickname"].ToString();
                realname = dt.Rows[0]["realname"].ToString();
            }
            else
            {
                nickname = "noNickname";
                realname = "noRealname";
            }
        }

        //update
            using (SqlConnection tSqlConn = new SqlConnection())
            {
                string webConfigConnectionString2 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["GardeningconnString"].ConnectionString;
                tSqlConn.ConnectionString = webConfigConnectionString2;
                tSqlConn.Open();

                string updateCommand = "update [user] set nickname = @nickname , display_name = @realname where user_id=@user_id ";
                SqlCommand updateCmd = new SqlCommand(updateCommand, tSqlConn);
                updateCmd.Parameters.AddWithValue("user_id", user);
                updateCmd.Parameters.AddWithValue("nickname", nickname);
                updateCmd.Parameters.AddWithValue("realname", realname);
                
                updateCmd.ExecuteNonQuery();
            }
    }
	private void LoadHotArticle()
    {
        string loadAscxPath = "~/Gardening/UserControls/HotArticle.ascx";
        Control control0 = Page.LoadControl(loadAscxPath);
        hot_PlaceHolder.Controls.Add(control0);
    }

    protected void PostArticle()
	{
	
		if (ViewState["hasLogin"] == null)
        {
            ViewState["hasLogin"] = WebUtility.CheckLogin(Session["memID"]).ToString();
        }
	    if (bool.Parse((string)ViewState["hasLogin"]))
        {
            if (gardeningService == null)
            {
                gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
            }
			Response.Write("<script language=javascript>window.location.href='/knowledge/knowledge_question.aspx?BackUrl=/knowledge/knowledge.aspx&gardening=true';</script>");
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "MyScript",
                "alert('請先登入會員');setHref('" + WebUtility.GetAppSetting("RedirectPage") + "');", true);
        }
    }
}
