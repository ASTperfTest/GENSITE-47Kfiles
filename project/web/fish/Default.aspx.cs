using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

public partial class _Default : Page
{
	string loginId="";
	string gameKey="";
	string realname = "";
	string nickname = "";
	string email = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        Login();
        FishBowl fb = new FishBowl();
		string ip_address = Request.UserHostAddress;
        fb.loginUpdate(loginId, gameKey, ip_address, nickname, realname, email);
        fb = null;
        GameStr.Text = createGame(loginId, gameKey);
    }

    // 產生遊戲內容
    private String createGame(string _login_id, string _game_key)
    {
        String Game = "";
        Game = "<script type='text/javascript'>\n";
        Game += "  var flashvars = {};\n";
        Game += "  var params = {};\n";
        Game += "  params.wmode = 'transparent';\n";
        Game += "  var attributes = {};\n";
        Game += "  swfobject.embedSWF('index.swf?login_id=" + _login_id + "&game_key=" + _game_key + "', 'main', '940', '570', '9.0.0', 'images/expressInstall.swf', flashvars, params, attributes);\n";
        Game += "</script>\n";
        return Game;
    }

    private void Login()
    {
        string guid = Request.QueryString["guid"];
        WebClient wc = new WebClient();

        //驗證此guid的資料
        wc.Encoding = UTF8Encoding.UTF8;
        string result = wc.DownloadString("http://kmweb.coa.gov.tw/sso/checklogin.aspx?guid=" + guid);

        //解析回傳的結果
        XmlDocument doc = new XmlDocument();
        doc.LoadXml(result);
        string isLogin = doc.GetElementsByTagName("IsLogin")[0].InnerText;
	
        //驗證成功則印出相關資訊,失敗則導回首頁
        if (isLogin == "true")
        {
            loginId = doc.GetElementsByTagName("UserId")[0].InnerText;
	    realname = doc.GetElementsByTagName("UserName")[0].InnerText;
	    nickname = doc.GetElementsByTagName("NickName")[0].InnerText;
	    email = doc.GetElementsByTagName("Email")[0].InnerText;	
            gameKey = guid;
        }
        else
        {
            Response.Write("<script>alert('驗證失敗，請重新登入!');location.href='/';</script>");
        }
    }
}
