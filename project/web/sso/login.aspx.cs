using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

public partial class sso_login : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
		Login();
    }

    private void Login()
    { 
        string guid = Request.QueryString["guid"];        
        WebClient wc = new WebClient(); 
		
        //驗證此guid的資料
		wc.Encoding = UTF8Encoding.UTF8;
        string result = wc.DownloadString("http://kmweb.coa.gov.tw/sso/checklogin.aspx?guid=" + guid );
		
		//解析回傳的結果
		XmlDocument doc = new XmlDocument();
        doc.LoadXml(result);
		string isLogin = doc.GetElementsByTagName("IsLogin")[0].InnerText;

        //驗證成功則印出相關資訊,失敗則導回首頁
		if (isLogin == "true")
        {
            foreach (XmlNode item in doc.GetElementsByTagName("Root")[0].ChildNodes)
            {
                Response.Write(item.Name + " : " + item.InnerText + "<br/>");
            }    
        }
		else 
        {
            Response.Write("<script>alert('驗證失敗，請聯絡系統管理員!');location.href='/';</script>");
        }
    }
}
