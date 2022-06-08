using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;

public partial class sso_checklogin : System.Web.UI.Page
{
    
    protected void Page_Load(object sender, EventArgs e)
    {
        CheckUserInfo();		
    }

    //檢查資料，並回傳xml檔給login.aspx
    private void CheckUserInfo()
    {
        string guId = Request.QueryString["guId"];
        string memId =  "";
        bool checkOK = false;
        string connString = System.Web.Configuration.WebConfigurationManager.AppSettings["COAConnectionString"];

        //用guid去SSO找CREATION_DATETIME欄位，比對現在時間，是否超過10秒
        string Command = @"SELECT * FROM SSO
									  WHERE GUID = @GUID";

        using (SqlConnection sqlConnObj = new SqlConnection(connString))
        {
            //設定查詢語法
			sqlConnObj.Open();
            SqlCommand tCmd = new SqlCommand(Command, sqlConnObj);
            tCmd.Parameters.AddWithValue("GUID", guId);
			
            //抓回資料
			SqlDataAdapter da = new SqlDataAdapter(tCmd);
            DataTable dt = new DataTable();
            da.Fill(dt);

            if (dt.Rows.Count != 0)
			{
				//取出驗證所需資訊			
				DateTime loginDateTime = Convert.ToDateTime(dt.Rows[0][2]);
				memId = Convert.ToString(dt.Rows[0][1]);

				//取出預設秒數
				double loginTime = Convert.ToDouble( System.Web.Configuration.WebConfigurationManager.AppSettings["LoginTime"]);				 
				//若登入驗證超過10秒則驗證失敗           
				if ( DateTime.Now < loginDateTime.AddSeconds(loginTime))
				{
					checkOK = true;
				}
				else
				{
					checkOK = false;
				}
			}
			else
			{
				//沒有資料也是驗證失敗
				checkOK = false;
			}
        }
        		
        Response.Expires = 0;
        Response.CacheControl = "no-cache";
        Response.AddHeader("Pragma", "no-cache");
		
        //驗證成功
        //用memId去Member找ACCOUNT，取出所需資訊，組XML，傳回給login.aspx解析
        //驗證失敗
        //僅傳部分資訊回給login.aspx解析
        if (checkOK)
        {
            string selCommand = @"SELECT * FROM MEMBER
                                    WHERE ACCOUNT = @ACCOUNT";
        
            using (SqlConnection sqlConnObj = new SqlConnection(connString))
            {
                //設定查詢語法
				sqlConnObj.Open();
                SqlCommand tCmd = new SqlCommand(selCommand, sqlConnObj);
                tCmd.Parameters.AddWithValue("ACCOUNT", memId);

                //抓出會員資料
				SqlDataAdapter da = new SqlDataAdapter(tCmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                if (dt.Rows.Count != 0)
                {
                    //驗證成功, 也找到會員資料時
					Response.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
                                    + "<Root>\n\t"
                                    + "<IsLogin>true</IsLogin>\n\t"
									+ "<Message>Login success!</Message>\n\t"
                                    + "<UserId>" + memId + "</UserId>\n\t"
                                    + "<UserName>" + HttpUtility.HtmlDecode(Convert.ToString(dt.Rows[0]["realname"]).Trim()) + "</UserName>\n\t"
                                    + "<Email>" + Convert.ToString(dt.Rows[0]["email"]).Trim() + "</Email>\n\t"
                                    + "<NickName>" + HttpUtility.HtmlDecode(Convert.ToString(dt.Rows[0]["nickname"]).Trim()) + "</NickName>\n"
                                    + "</Root>");
                }
                else
                {
                    //驗證成功, 但找不到會員資料時
					Response.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
                            + "<Root>\n\t"
							+ "<IsLogin>false</IsLogin>\n\t"
							+ "<Message>Member not found!</Message>\n\t"
							+ "</Root>");
                }
            }
        }
        else
        {
            //驗證失敗時			
			Response.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n"
                            + "<Root>\n\t"
							+ "<IsLogin>false</IsLogin>\n\t"
							+ "<Message>Login faild!</Message>\n\t"
							+ "</Root>");
        }
    }
}
