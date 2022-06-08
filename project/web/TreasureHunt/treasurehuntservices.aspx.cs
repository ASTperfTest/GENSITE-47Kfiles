using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Collections.Specialized;
using System.Collections.Generic;

public partial class TreasureHunt_treasurehuntservices : System.Web.UI.Page
{
    private string loginId = "";
    private TreasureHunt treasureHunt;
    string connString = System.Web.Configuration.WebConfigurationManager.AppSettings["COAConnectionString"];
    protected void Page_Load(object sender, EventArgs e)
    {
        treasureHunt = new TreasureHunt("");
        string cmd = WebUtility.GetStringParameter("cmd", string.Empty).ToLower();
        string targetpage = WebUtility.GetStringParameter("targetpage", string.Empty).ToLower();
        string pageParam = WebUtility.GetStringParameter("documentinfo", string.Empty).ToLower();
        string referer_url = WebUtility.GetStringParameter("refereurl", string.Empty).ToLower();
        if (Session[pageParam] != null && Session["memID"] != null)
        {
            Dictionary<string, string> temp = Session[pageParam] as Dictionary<string, string>;
            if (string.Compare(temp["pageParam"].ToString(), targetpage + pageParam)==0)
                referer_url = temp["referer_url"].ToString();
        }
        if (!treasureHunt.checkActivity(Server.UrlDecode((targetpage)), Server.UrlDecode(pageParam), Server.UrlDecode(referer_url)))
            return;
        if (Session["memID"] == null || Session["memID"].ToString().Trim() == "")
        {
            if (cmd.CompareTo("checkmembrtcaptchat") == 0)
            {
                GetTimeOutMessage();
                return;
            }
            if (Session["GuestNegateTreasure"] != null && (Session["GuestNegateTreasure"].ToString().Trim() != ""))
                return;
            if (cmd.CompareTo("setguestnegatetreasure") == 0)
                SetGuestNegateTreasure();
            ReadDocument readDocument = new ReadDocument();
            readDocument.Read = false;
            readDocument.Message = HttpUtility.UrlEncode("您無法進行尋寶活動，這可能因為您尚未登錄會員，請登錄後再嘗試訪問頁面。");
            readDocument.Image = "21.jpg";
            readDocument.Html = HttpUtility.UrlEncode("<br/><div class=\"content\" style=\"padding-left:5px;text-align:left;\"><input type=\"checkbox\" id=\"notallow\" onclick=\"SetGuestNegateTreasure()\" value=\"\" />不要再出現此訊息<div>");
            readDocument.RefererUrl = HttpUtility.UrlEncode(referer_url);
            Dictionary<string, string> temp = new Dictionary<string, string>();
            temp.Add("referer_url", referer_url);
            temp.Add("pageParam", targetpage + pageParam);
            Session[pageParam] = temp;
            WebUtility.WriteAjaxResult(true, "", readDocument);
            return;
        }
        else
        {
            loginId = Session["memID"].ToString().Trim();
            if (Session["LoginUserNegateTreasure"] != null && Session["LoginUserNegateTreasure"].ToString().CompareTo(loginId) == 0)
                return;
            treasureHunt = new TreasureHunt(loginId);
             pageParam = HttpUtility.HtmlDecode(pageParam);
             if (treasureHunt.getDocumentId(targetpage + pageParam) == -1)
                 return;
            string selCommand = " SELECT * FROM MEMBER WHERE ACCOUNT = @ACCOUNT ";
            using (SqlConnection sqlConnObj = new SqlConnection(connString))
            {
                //設定查詢語法
                sqlConnObj.Open();
                SqlCommand tCmd = new SqlCommand(selCommand, sqlConnObj);
                tCmd.Parameters.AddWithValue("ACCOUNT", loginId);

                //抓出會員資料
                SqlDataAdapter da = new SqlDataAdapter(tCmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                //檢查遊戲使用者及更新
                if (dt.Rows.Count != 0)
                {
                    if (Session["treasurehunt"] == null || Session["treasurehunt"].ToString().Trim() == "")
                    {
                        treasureHunt.loginUpdate( HttpUtility.HtmlDecode(Convert.ToString(dt.Rows[0]["nickname"]).Trim()),
                             HttpUtility.HtmlDecode(Convert.ToString(dt.Rows[0]["realname"]).Trim()),Convert.ToString(dt.Rows[0]["email"]).Trim());
                    }
                }
                else
                {
                    return;
                }
            }
        }
        if (cmd == string.Empty)
        {
            return;
        }
        else
        {
            DispatchMethod(cmd);
        }
    }




    private void DispatchMethod(string MethodName)
    {
        switch (MethodName)
        {
            case "checkreaddocument":
                checkReadDocument();
                break;
            case "searchtreasurebox":
                SearchTreasureBox();
                break;
            case "probablity":
                GEtprobablity();
                break;
            case "checkmembrtcaptchat":
                CheckMembrtCapthat();
                break;
            case "getmembercaptchat":
                GetMemberCaptchat();
                    break;
            case "setloginusernegatetreasure":
                    SetLoginUserNegateTreasure();
                    break;
            default:
                break;
        }
    }

    private void GEtprobablity()
    {
        //string targetpage = WebUtility.GetStringParameter("targetpage", string.Empty).ToLower();
        //string pageParam = WebUtility.GetStringParameter("documentinfo", string.Empty).ToLower();
        //WebUtility.WriteAjaxResult(true, "", treasureHunt.getValueee());
        string sql = " {0}  ";
        string temp = "";
        int q = 3;
        for (int i = 0; i < q; i++)
        {
            temp = " IF EXISTS( ";
            temp += " SELECT * FROM READ_DOCUMENT_RECORD ";
            temp += " WHERE ACTIVITY_ID = 1 AND ACCOUNT_ID = 7 AND URL = '/subject/ct.asp?xitem=156311&ctnode=4147&mp=86&kpi=0' AND  DateAdd(\"d\","+(q-i).ToString()+",MODIFY_DATE) >  GETDATE() )";
            temp += " {0} ";
            temp += " ELSE SELECT " + i.ToString() + "  ";
            sql = string.Format(sql,temp);
        }
        sql = string.Format(sql,"SELECT "+q.ToString());
        WebUtility.WriteAjaxResult(true, "", sql);
    }

    private void checkReadDocument()
    {
        string targetpage = WebUtility.GetStringParameter("targetpage", string.Empty).ToLower();
        string pageParam = WebUtility.GetStringParameter("documentinfo", string.Empty).ToLower();
        string referer_url = WebUtility.GetStringParameter("refereurl", string.Empty).ToLower();
        if (Session[pageParam] != null)
        {
            Dictionary<string, string> temp = Session[pageParam] as Dictionary<string, string>;
            if (string.Compare(temp["pageParam"].ToString(), targetpage + pageParam) == 0)
            {
                referer_url = temp["referer_url"].ToString();
            }
        }
        pageParam = HttpUtility.HtmlDecode(pageParam);
        targetpage = HttpUtility.HtmlDecode(targetpage);

        if (targetpage == "" || !treasureHunt.checkActivity(targetpage, pageParam, referer_url))
            return;

        
        ReadDocument readDocument = new ReadDocument();
        if (treasureHunt.CheckDailyGetLimit())
        {
            readDocument.Read = false;
            readDocument.Image = "41.jpg";
            readDocument.Message = HttpUtility.UrlEncode("會員一日至多可得到三樣寶物<br/>（不分種類，您今日已達到<br/>寶物蒐集上限，歡迎明天繼續<br/>挑戰</li>");
            readDocument.Html = HttpUtility.UrlEncode("<div class=\"content\" style=\"padding-left:5px;text-align:left;\"><br/><input type=\"checkbox\" id=\"notallow\" onclick=\"SetLoginUserNegateTreasure()\" value=\"\" />不要再出現此訊息</div>");
            WebUtility.WriteAjaxResult(true, "", readDocument);
            return;
        }
        int intervallDays = treasureHunt.checkReadDocument(targetpage + pageParam);
        if (intervallDays == 0 && treasureHunt.CheckLastSearchTime())
        {
            readDocument.Read = true;
            readDocument.Reciprocalime = treasureHunt.reciprocal_time;
            readDocument.Message = "尋寶遊戲還在倒數中，您確定要離開這個瀏覽頁?";
            readDocument.Image = "61.jpg";
            if (Session[pageParam] != null)
            {
                Dictionary<string, string> temp = Session[pageParam] as Dictionary<string, string>;
                if (string.Compare(temp["pageParam"].ToString(), targetpage + pageParam) == 0)
                {
                    Session.Remove(pageParam);
                    readDocument.RefererUrl = HttpUtility.UrlEncode(referer_url);
                }
            }
            Dictionary<string, string> userData = treasureHunt.GetUserData(loginId);
            WebUtility.WriteAjaxResult(true, "", readDocument);
        }
        else if (intervallDays == 0 && !treasureHunt.CheckLastSearchTime())
        {
            readDocument.Read = false;
            readDocument.Message = HttpUtility.UrlEncode("同一使用者同時間僅能開啟一個有效尋寶頁面，此頁面無法進行尋寶活動之倒數計時。請留意有效網頁，並於該頁面進行尋寶/獲寶的動作喔！");
            readDocument.Image = "11.jpg";
            WebUtility.WriteAjaxResult(true, "" , readDocument);
        }
        else
        {
            readDocument.Read = false;
            readDocument.Message = HttpUtility.UrlEncode(ReadMessage(intervallDays));
            readDocument.Image = ReadImage(intervallDays);
            WebUtility.WriteAjaxResult(true, "", readDocument);
        }
    }

    private string ReadMessage(int days)
    {
        string message = "";
        switch (days)
        {
            case 1:
                message = "尋寶頁面五日內僅可尋寶一次;本頁面已於四天前尋寶完畢，歡迎明天再度光臨，謝謝";
                break;

            case 2:
                message = "尋寶頁面五日內僅可尋寶一次;本頁面已於三天前尋寶完畢，歡迎後天再度光臨，謝謝";
                break;
            case 3:
                message = "尋寶頁面五日內僅可尋寶一次;本頁面已於前日尋寶完畢，歡迎三天後再度光臨，謝謝";
                break;
            case 4:
                message = "尋寶頁面五日內僅可尋寶一次;本頁面已於昨日尋寶完畢，歡迎四天後再度光臨，謝謝";
                break;
            case 5:
                message = "尋寶頁面五日內僅可尋寶一次;本頁面已於今日尋寶完畢，歡迎五天後再度光臨，謝謝";
                break;
            default:
                break;
        }
        return message;
    }

    private string ReadImage(int days)
    {
        string message = "";
        switch (days)
        {
            case 1:
                message = "53.jpg";
                break;

            case 2:
                message = "52.jpg";
                break;
            case 3:
                message = "51.jpg";
                break;
            case 4:
                message = "50.jpg";
                break;
            case 5:
                message = "49.jpg";
                break;
            default:
                break;
        }
        return message;
    }
    
    private void SearchTreasureBox()
    {
        string targetpage = WebUtility.GetStringParameter("targetpage", string.Empty).ToLower();
        string pageParam = WebUtility.GetStringParameter("documentinfo", string.Empty).ToLower();
        pageParam = HttpUtility.HtmlDecode(pageParam);
        targetpage = HttpUtility.HtmlDecode(targetpage);
        string referer_url = WebUtility.GetStringParameter("refereurl", string.Empty).ToLower();
        if (targetpage == "" || !treasureHunt.checkActivity(targetpage, pageParam, HttpUtility.HtmlDecode(referer_url)))
            return;
        bool toUpperLimit = treasureHunt.CheckDailyGetLimit();
        int intervallDays = treasureHunt.checkReadDocument(targetpage + pageParam);
        if (intervallDays == 0 && !toUpperLimit)
        {
            bool find = treasureHunt.SearchTreasureBox(targetpage + pageParam);
            int treasureLofId = treasureHunt.SaveReadDocimentRecordData(targetpage + pageParam, find);
            if (find)
            {
				Session["treasureLogId"] = treasureLofId;
                ReadDocument readDocument = new ReadDocument();
                readDocument.Read = true;
                readDocument.TreasureLodId = treasureLofId;
                WebUtility.WriteAjaxResult(true, "", readDocument);
            }
            else
            {
                ReadDocument readDocument = new ReadDocument();
                readDocument.Read = false;
                readDocument.Message = HttpUtility.UrlEncode("寶物不在這邊，請到其他頁面<br/>繼續尋寶，下次尋獲寶物的機<br/>率將提昇，祝您好運！");
                readDocument.Image = "62.jpg";
                WebUtility.WriteAjaxResult(true, "", readDocument);
            }
        }
        else if (toUpperLimit)
        {
            ReadDocument readDocument = new ReadDocument();
            readDocument.Read = false;
            readDocument.Image = "41.jpg";
            readDocument.Message = HttpUtility.UrlEncode("會員一日至多可得到三樣寶物<br/>（不分種類，您今日已達到<br/>寶物蒐集上限，歡迎明天繼續<br/>挑戰</li>");
            readDocument.Html = HttpUtility.UrlEncode("<div class=\"content\" style=\"padding-left:5px;text-align:left;\"><br/><input type=\"checkbox\" id=\"notallow\" onclick=\"SetLoginUserNegateTreasure()\" value=\"\" />不要再出現此訊息</div>");
            WebUtility.WriteAjaxResult(true, "", readDocument);
        }
        else
        {
            WebUtility.WriteAjaxResult(true, HttpUtility.UrlEncode(ReadMessage(intervallDays)), false);
        }
    }

    private void CheckMembrtCapthat()
    {
        string targetpage = WebUtility.GetStringParameter("targetpage", string.Empty).ToLower();
        string pageParam = WebUtility.GetStringParameter("documentinfo", string.Empty).ToLower();
        string codeString = WebUtility.GetStringParameter("code", string.Empty).ToLower();
        string treasureLogIdString = WebUtility.GetStringParameter("treasurelodid", string.Empty).ToLower();
        string guid = WebUtility.GetStringParameter("guid", string.Empty).ToLower();
        int treasureLogId = -1;
        pageParam = HttpUtility.HtmlDecode(pageParam);
        targetpage = HttpUtility.HtmlDecode(targetpage);
        string referer_url = WebUtility.GetStringParameter("refereurl", string.Empty).ToLower();
        if (!string.IsNullOrEmpty(treasureLogIdString))
            treasureLogId = Convert.ToInt32(treasureLogIdString);
        if (targetpage == "" || !treasureHunt.checkActivity(targetpage, pageParam, HttpUtility.UrlDecode(referer_url)))
            return;
        if (Session[guid+"CaptchaImageText"] == null || string.IsNullOrEmpty(Session[guid+"CaptchaImageText"].ToString()) ||  string.IsNullOrEmpty(codeString) ||
            string.Compare(Session[guid+"CaptchaImageText"].ToString(), codeString, false) != 0)
        {
            WebUtility.WriteAjaxResult(true, HttpUtility.UrlEncode("驗證碼輸入錯誤!請重新輸入"), false);
        }
        else
        {
            Session.Remove(guid + "CaptchaImageText");
            bool toUpperLimit = treasureHunt.CheckDailyGetLimit();
            if (!toUpperLimit)
            {
                int  treasureId = treasureHunt.UpdateTreasureHuntLog(treasureLogId);
                if (treasureId !=0)
                {
                    TreasureHunt.Treasure treasure = treasureHunt.GetTreasureById(treasureId,-1);
                    ReadDocument readDocument = new ReadDocument();
                    readDocument.Read = true;
                    readDocument.TreasureName = HttpUtility.UrlEncode(treasure.TreasureName);
                    readDocument.Image = treasure.Icon;
                    WebUtility.WriteAjaxResult(true,"", readDocument);
                }
                else
                {
                    ReadDocument readDocument = new ReadDocument();
					readDocument.Read = false;
					readDocument.Image = "62.jpg";
					readDocument.Message = HttpUtility.UrlEncode("請確認尋寶途徑是否正確!!");
					readDocument.Html = HttpUtility.UrlEncode("<div class=\"content\" style=\"padding-left:5px;text-align:left;\"><br/><input type=\"checkbox\" id=\"notallow\" onclick=\"SetLoginUserNegateTreasure()\" value=\"\" />不要再出現此訊息</div>");
					WebUtility.WriteAjaxResult(true, "", readDocument);
                }
            }
            else
            {
                ReadDocument readDocument = new ReadDocument();
                readDocument.Read = false;
                readDocument.Image = "41.jpg";
                readDocument.Message = HttpUtility.UrlEncode("會員一日至多可得到三樣寶物<br/>（不分種類，您今日已達到<br/>寶物蒐集上限，歡迎明天繼續<br/>挑戰</li>");
                readDocument.Html = HttpUtility.UrlEncode("<div class=\"content\" style=\"padding-left:5px;text-align:left;\"><br/><input type=\"checkbox\" id=\"notallow\" onclick=\"SetLoginUserNegateTreasure()\" value=\"\" />不要再出現此訊息</div>");
                WebUtility.WriteAjaxResult(true, "", readDocument);
            }
        }
        
    }


    private void SetGuestNegateTreasure()
    {
        Session["GuestNegateTreasure"] = "true";
    }

    private void SetLoginUserNegateTreasure()
    {
        Session["LoginUserNegateTreasure"] = Session["memID"].ToString().Trim();
    }

    private void GetMemberCaptchat()
    {
        string targetpage = WebUtility.GetStringParameter("targetpage", string.Empty).ToLower();
        string pageParam = WebUtility.GetStringParameter("documentinfo", string.Empty).ToLower();
        string codeString = WebUtility.GetStringParameter("code", string.Empty).ToLower();
        string treasureLogIdString = WebUtility.GetStringParameter("treasurelodid", string.Empty).ToLower();
        WebUtility.WriteAjaxResult(true, HttpUtility.UrlEncode("恭喜取得寶物"), "");
    }

    private void GetTimeOutMessage()
    {
        ReadDocument readDocument = new ReadDocument();
        readDocument.Read = false;
        readDocument.Message = HttpUtility.UrlEncode("您閒置超過20分鐘，請重新登入以再次尋寶。");
        readDocument.Timeout = true;
        readDocument.Image = "21.jpg";
        readDocument.Html = HttpUtility.UrlEncode("<br/><div class=\"content\" style=\"padding-left:5px;text-align:left;\"><input type=\"checkbox\" id=\"notallow\" onclick=\"SetGuestNegateTreasure()\" value=\"\" />不要再出現此訊息<div>");
        WebUtility.WriteAjaxResult(true, "", readDocument);
    }

    class ReadDocument 
    {
        bool m_Read;
        bool m_Timeout;
        int m_TreasureLodId;
        string m_TreasureName;
        string m_UserName;
        string m_Reciprocalime;
        string m_Message;
        string m_Html;
        string m_Image;
        string m_Message2;
        string m_Message3;
        string m_Message4;
        string m_Message5;
        string m_RefererUrl;

        public bool Read
        {
            get { return m_Read; }
            set { m_Read = value; }
        }
        public bool Timeout
        {
            get { return m_Timeout; }
            set { m_Timeout = value; }
        }
        public int TreasureLodId
        {
            get { return m_TreasureLodId; }
            set { m_TreasureLodId = value; }
        }
        public string TreasureName
        {
            get { return m_TreasureName; }
            set { m_TreasureName = value ;}
        }
        public string UserName
        {
            get { return m_UserName; }
            set { m_UserName = value; }
        }
        public string Message
        {
            get { return m_Message; }
            set { m_Message = value; }
        }
        public string Message2
        {
            get { return m_Message2; }
            set { m_Message2 = value; }
        }
        public string Message3
        {
            get { return m_Message3; }
            set { m_Message3 = value; }
        }
        public string Message4
        {
            get { return m_Message4; }
            set { m_Message4 = value; }
        }
        public string Message5
        {
            get { return m_Message5; }
            set { m_Message5 = value; }
        }
        public string Reciprocalime
        {
            get { return m_Reciprocalime; }
            set { m_Reciprocalime = value; }
        }
        public string Html
        {
            get { return m_Html; }
            set { m_Html = value; }
        }
        public string Image
        {
            get { return m_Image; }
            set { m_Image = value; }
        }
        public string RefererUrl
        {
            get { return m_RefererUrl; }
            set { m_RefererUrl = value; }
        }
    }
}
