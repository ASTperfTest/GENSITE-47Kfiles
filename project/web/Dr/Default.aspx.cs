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
        if (IsActivityDuration())
        {
            Login();

            // 更新會員key
            Farm2009 farm = new Farm2009(loginId, gameKey);

            farm.loginUpdate(nickname, realname, email);
            farm = null;
            GameStr.Text = createGame(loginId, gameKey);
        }
        else {
            GameStr.Text = "<fieldset style='width:60%;background-color:#E2FFE1;'><legend>活動資訊</legend>2010全能農知博識王的活動已到期(2010 8/13-8/23 早上十點)<br>遊戲將暫停開放，待抽獎完成後再重新開啟<br>前往<a href='http://kmweb.coa.gov.tw/dr/top_20102.aspx'>活動排行榜</a> 或是<a href='http://kmweb.coa.gov.tw/'>回到入口網</a></fieldset>";
        }
    }
	
	private bool IsActivityDuration()
	{
        if (DateTime.Compare(DateTime.Now, new DateTime(2010, 8, 23, 10, 0, 0)) >= 0)
        {
            return false;
        }
        else {
            return true;
        }
	}
	
 
    // 產生遊戲內容

    private String createGame(string _login_id, string _game_key)
    {
        String Game = "";
        Game = "<script type='text/javascript'>\n";
        Game += "  var flashvars = {};\n";
        Game += "  var params = {};\n";
        //Game += "  params.wmode = 'transparent';\n";
        Game += "  var attributes = {};\n";
        Game += "  swfobject.embedSWF('index.swf?login_id=" + _login_id + "&game_key=" + _game_key + "', 'main', '900', '642', '9.0.0', 'images/expressInstall.swf', flashvars, params, attributes);\n";
        Game += "</script>\n";
        return Game;
    }
 
    private void Login()
    {
        string guid = Request.QueryString["guid"];

        WebClient wc = new WebClient();
 
        //驗證此guid的資料
        wc.Encoding = UTF8Encoding.UTF8;
        //string result = wc.DownloadString("http://kwpi-coa-kmweb.gss.com.tw/sso/checklogin.aspx?guid=" + guid);
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
 