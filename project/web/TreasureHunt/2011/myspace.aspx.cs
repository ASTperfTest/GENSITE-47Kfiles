using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using log4net;
using System.Collections.Generic;

public partial class TreasureHunt_2011_myspace : System.Web.UI.Page
{
    private ILog log4;
    private TreasureHunt treasureHunt;
    private int activityid = 2;
    string connString = System.Web.Configuration.WebConfigurationManager.AppSettings["COAConnectionString"];
    protected string buttonDisable="true";
    protected string voteDisplay = "false";
    protected string usersTreasureJson = "";
    protected string TreasureDropDownListClientId = "";
    protected string TreasureDropDownList2ClientId = "";
    protected string changeTreasureBox = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        log4 = LogManager.GetLogger("DR");
        if (Session["memID"] == null || Session["memID"].ToString().Trim() == "")
        {
            if (!(WebUtility.GetStringParameter("type", string.Empty).CompareTo("logout") == 0))
            {
                Response.Write("<script>alert('連線逾時或尚未登入，請登入會員!');history.go(-1);</script>");
            }
            else
            {
                Response.Write("<script>location.href=\"/mp.asp?mp=1\";</script>");
            }
            return;
        }
        activityid = (WebUtility.GetStringParameter("activityid", string.Empty) == "") ? 3 : Convert.ToInt32(WebUtility.GetStringParameter("activityid", "3"));
        if (activityid != 3)
        {
            Response.Redirect("myspace.aspx?activityid=3");
            Response.End();
        }
        string loginId = Session["memID"].ToString().Trim();
        if (SaveUserData(loginId) == 0)
        {
            Response.Write("<script>alert('連線逾時或尚未登入，請登入會員!');history.go(-1);</script>");
            return;
        }
        treasureHunt = new TreasureHunt(loginId);
        treasureHunt.SetActivity(activityid);
        randomChangeTreasure.Attributes.Add("onclick", "return check();");
        if (!string.IsNullOrEmpty(Request.Form["__EVENTTARGET"]))
            if (Request.Form["__EVENTTARGET"].Replace("$", "_") == randomChangeTreasure.ClientID)
            {
                if (TreasureDropDownList.SelectedIndex != 0)
                    RanDomChangeTreasure();
            }
        if (treasureHunt.VoteLotteryIsEnable())
        {
            voteDisplay = "true";
            if (Request.Form.GetValues("hideVote") != null && !string.IsNullOrEmpty(Request.Form.GetValues("hideVote")[0].ToString()))
                SetVoteLottery(Session["memID"].ToString().Trim());
        }
        SetUserTreasure();
        SetUserDetal(Session["memID"].ToString().Trim());
        Session.Remove("voteLottery");
    }

    private void SetUserDetal(string accountId)
    {
        Dictionary<string, string> userData = treasureHunt.GetUserData(accountId);
        if (userData["loginId"] != null)
            loginid.Text = userData["loginId"];
        if (userData["name"] != null)
            userName.Text = userData["name"];
        int totalsuit = treasureHunt.GetUsersTotalSuit(accountId);
        UsersTotalSuit.Text = totalsuit.ToString();
        IList lotteryList = treasureHunt.GetUserVotesLottery(activityid, treasureHunt.account_id);
        int UseForLottery = 0;
        foreach (TreasureHunt.LotteryGift temp in lotteryList)
        {
            UseForLottery += temp.Votes;
            SetLabelCount(temp.GiftId,temp.Votes);
        }
        Lottery.Text = UseForLottery.ToString();
        if ((totalsuit - UseForLottery) > 0)
            buttonDisable = "false";
        BeforeLottery.Text = (totalsuit - UseForLottery).ToString();
    }

    private int SaveUserData(string loginId)
    {
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
                treasureHunt = new TreasureHunt(loginId);
                if (Session["treasurehunt"] == null || Session["treasurehunt"].ToString().Trim() == "")
                {
                    treasureHunt.loginUpdate(HttpUtility.HtmlDecode(Convert.ToString(dt.Rows[0]["nickname"]).Trim()),
                         HttpUtility.HtmlDecode(Convert.ToString(dt.Rows[0]["realname"]).Trim()), Convert.ToString(dt.Rows[0]["email"]).Trim());
                }
            }
            else
            {
                return 0;
            }
        }
        return 1;
    }

    private void SetUserTreasure()
    {
        string loginId = Session["memID"].ToString().Trim();
        string ss = "";
        Dictionary<int, TreasureHunt.Treasure> treasureDetal = treasureHunt.GetTreasureDetal();
        Dictionary<int, int> usersTreasure = treasureHunt.GetUsersTreasure(treasureHunt.account_id);
        int maxSiut = treasureHunt.limit_set;
        ss = "<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" id=\"reset\">";
        Dictionary<int, TreasureHunt.Treasure>.KeyCollection treasureKey = treasureDetal.Keys;
        IList treasureList = new ArrayList();

        bool canChangeTreasure = false;
        canChangeTreasure = treasureHunt.CanChangeTreasure(treasureHunt.account_id, treasureHunt.getactivity_id());
        canChangeTreasure = true;
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;
            treasureList.Add(treasure);
        }
        int i = 0;
        int maxPieces = 0;
        while(i<treasureList.Count)
        {
            TreasureHunt.Treasure treasure = new TreasureHunt.Treasure();
            ss += "<tr>";
            int temp = 0;
            int changeId = 0;
            if (i < treasureList.Count)
            {
                treasure = (TreasureHunt.Treasure)treasureList[i];
                temp = GetUserTrasureCount(usersTreasure, treasure.TreasureId);
                maxPieces += temp;
                if (temp == 0)
                {
                    treasure.Icon = "bean_q.gif";
                    changeId = 0;
                }
                else
                {
                    changeId = treasure.TreasureId;
                }
                if (canChangeTreasure)
                {
                    ss += "<td width=\"20%\" align=\"center\"><div id=\"trediv" + treasure.TreasureId + "\" style=\"width:90px;\" name=\"changetreasure\" onclick=\"changeTreasure(" + changeId + ",'" + treasure.Icon + "',this);\"><img src=\"image/" + treasure.Icon + "\" /></div><br />";
                }
                else
                {
                    ss += "<td width=\"20%\" align=\"center\"><img src=\"image/" + treasure.Icon + "\" /><br />";
                }
                ss += "X" + temp + "</td>";
                i = i + 1;
            }
            if (i < treasureList.Count)
            {
                treasure = (TreasureHunt.Treasure)treasureList[i];
                temp = GetUserTrasureCount(usersTreasure, treasure.TreasureId);
                maxPieces += temp;
                if (temp == 0)
                {
                    treasure.Icon = "bean_q.gif";
                    changeId = 0;
                }
                else
                {
                    changeId = treasure.TreasureId;
                }
                if (canChangeTreasure)
                {
                    ss += "<td width=\"20%\" align=\"center\"><div id=\"trediv" + treasure.TreasureId + "\" style=\"width:90px;\"  name=\"changetreasure\" onclick=\"changeTreasure(" + changeId + ",'" + treasure.Icon + "',this);\"><img src=\"image/" + treasure.Icon + "\" /></div><br />";
                }
                else
                {
                    ss += "<td width=\"20%\" align=\"center\"><img src=\"image/" + treasure.Icon + "\" /><br />";
                } 
                ss += "X" + temp + "</td>";
                i = i + 1;
            }
            else
            {
                ss += "<td width=\"20%\" align=\"center\"></td>";
            }
            if (i < treasureList.Count)
            {

                treasure = (TreasureHunt.Treasure)treasureList[i];
                temp = GetUserTrasureCount(usersTreasure, treasure.TreasureId);
                maxPieces += temp;
                if (temp == 0)
                {
                    treasure.Icon = "bean_q.gif";
                    changeId = 0;
                }
                else
                {
                    changeId = treasure.TreasureId;
                }
                if (canChangeTreasure)
                {
                    ss += "<td width=\"20%\" align=\"center\"><div id=\"trediv" + treasure.TreasureId + "\" style=\"width:90px;\" name=\"changetreasure\" onclick=\"changeTreasure(" + changeId + ",'" + treasure.Icon + "',this);\"><img src=\"image/" + treasure.Icon + "\" /></div><br />";
                }
                else
                {
                    ss += "<td width=\"20%\" align=\"center\"><img src=\"image/" + treasure.Icon + "\" /><br />";
                }
                ss += "X" + temp + "</td>";
                i = i + 1;
            }
            else
            {
                if (canChangeTreasure)
                {
                    ss += "<td width=\"20%\" align=\"center\"><br/><br/><br/><br/><br/><br/><br/><input id=\"iwantChangei\" type=\"button\" value=\"我要交換\" onclick=\"IwantChange(this);\" /></td>";
                }
                else{
                    ss += "<td width=\"20%\" align=\"center\"></td>";     
                }
            }
            ss += "<tr>";
        }
        ss += "</table>";
        treasureTable.Text = ss;
        //todo
        //要看技正開放兌換的時間為何
        if (canChangeTreasure && maxPieces >= 2 && activityid == 3)
        {
            randomPanel.Visible = true;
            TreasureDropDownList.DataSource = CreateDataSource(treasureDetal, usersTreasure);
            TreasureDropDownList.DataTextField = "TextField";
            TreasureDropDownList.DataValueField = "ValueField";
            TreasureDropDownList.DataBind();
            TreasureDropDownListClientId = TreasureDropDownList.ClientID;
            TreasureDropDownList2.DataSource = CreateDataSource(treasureDetal, usersTreasure);
            TreasureDropDownList2.DataTextField = "TextField";
            TreasureDropDownList2.DataValueField = "ValueField";
            TreasureDropDownList2.DataBind();
            TreasureDropDownList2ClientId = TreasureDropDownList2.ClientID;
            usersTreasureJson = Server.UrlEncode(WebUtility.convertToJsonString(GetUserTreasure(treasureList, usersTreasure)));
            TreasureDropDownList.Visible = true;
            TreasureDropDownList2.Visible = true;
        }
        else if (activityid ==3)
        {
            if (maxPieces < 2)
            {
                canotchangeMessage.Text = "您沒有足夠的寶物可進行交換!!";
            }
            randomPanel.Visible = false;
            acs.Visible = true;
        }
        TreasureDropDownList.SelectedIndex = 0;
    }

    private int GetUserTrasureCount(Dictionary<int, int> usersTreasure,int id)
    {
        int returnString = 0;
        if (usersTreasure.ContainsKey(id))
        {

            returnString =usersTreasure[id];
        }

        return returnString;
    }

    private void SetVoteLottery(string accountId)
    {
        string giftId= Request.Form.GetValues("hideVote")[0].ToString();
        if (Session["voteLottery"] != null && Session["voteLottery"].ToString() == "true")
        {
            Response.Write("<script type=\"text/javascript\">alert(\"請勿重複投套數!\");</script>");
            Response.End();
        }
        Session["voteLottery"] = "true";
        int userTotalSuit = treasureHunt.GetUsersTotalSuit(accountId);
        IList lotteryList = treasureHunt.GetUserVotesLottery(activityid, treasureHunt.account_id);
        int UseForLottery = 0;
        foreach (TreasureHunt.LotteryGift temp in lotteryList)
        {
            UseForLottery += temp.Votes;
        }
        if (userTotalSuit <= UseForLottery)
        {
            Response.Write("<script type=\"text/javascript\">alert(\"您可兌換的套數已經沒了唷!\");</script>");
            return;
        }
        if (giftId != "D" && giftId != "E" && giftId != "F")
        {
            Response.Write("<script type=\"text/javascript\">alert(\"請確認您所要投的獎項是否正確!\");</script>");
            return;
        }
        treasureHunt.SetUserVotesLottery(activityid, treasureHunt.account_id, giftId);
    }
    private void SetLabelCount(string giftId,int count)
    {
        switch (giftId)
        {
            case "D":
                labelLotteryA.Text = count.ToString();
                break;
            case "E":
                labelLotteryB.Text = count.ToString();
                break;
            case "F":
                labelLotteryC.Text = count.ToString();
                break;
            default:
                break;
        }
    }

    protected void RanDomChangeTreasure()
    {
        if (Session["voteLottery"] != null && Session["voteLottery"].ToString() == "true")
        {
            Response.Write("<script type=\"text/javascript\">alert(\"請勿重複兌換!\");</script>");
            Response.Redirect("/");
            Response.End();
        }
        bool canchange =  treasureHunt.CanChangeTreasure(treasureHunt.account_id, treasureHunt.getactivity_id());
        canchange = true;
        if (!canchange)
        {
            Response.Write("<script type=\"text/javascript\">alert(\"請勿重複兌換!\");</script>");
            Response.Redirect("/");
            Response.End();
        }
        Session["voteLottery"] = "true";
        int changeTreasureId = 0;
        int changeTreasureId2 = 0;
        int getTreasureID = 0;
        if (int.TryParse(TreasureDropDownList.SelectedValue, out changeTreasureId) && int.TryParse(TreasureDropDownList2.SelectedValue, out changeTreasureId2))
        {
            getTreasureID = treasureHunt.RandomChangeTreasure(treasureHunt.account_id, treasureHunt.getactivity_id(), changeTreasureId, changeTreasureId2);
            switch (getTreasureID)
            {
                case 0:
                    Response.Write("<script type=\"text/javascript\">alert(\"請確認您要拿來交換的寶物是否有兩個以上!\");window.location=\"myspace.aspx?activityid=" + activityid.ToString()+ "\"</script>");
                    Response.End();
                    break;
                case -1:
                    Response.Write("<script type=\"text/javascript\">alert(\"伺服器發生錯誤請稍後在試!\");window.location=\"myspace.aspx?activityid=" + activityid.ToString()+ "\"</script>");
                    Response.End();   
                    break;
                case -2:
                    Response.Write("<script type=\"text/javascript\">alert(\"您今天換過寶物了喔!\");window.location=\"myspace.aspx?activityid=" + activityid.ToString() + "\"</script>");
                    Response.End();
                    break;
                default:
                    break;
            }
            //TreasureHunt.Treasure treasure = treasureHunt.GetTreasureById(getTreasureID, activityid);
            Session["changeTreasureId"] = getTreasureID.ToString();
            changeTreasureBox = "/treasurehunt/2011/change.aspx";
            treasureall.Visible = false;
            changebox.Visible = true;
        }
    }

    private Dictionary<int, User_Treasure> GetUserTreasure(IList treasureList, Dictionary<int, int> userTreasure)
    {
        Dictionary<int, User_Treasure> returnlist = new Dictionary<int, User_Treasure>();
       
        foreach (TreasureHunt.Treasure tr in treasureList)
        {
            User_Treasure ut = new User_Treasure();
            ut.TreasureId = tr.TreasureId;
            ut.TreasureName = tr.TreasureName;
            ut.Piece = 0;
            if (userTreasure.ContainsKey(tr.TreasureId))
            {
                ut.Piece = userTreasure[tr.TreasureId];
            }
            returnlist.Add(tr.TreasureId, ut);
        }

        return returnlist;
    }

    ICollection CreateDataSource(Dictionary<int, TreasureHunt.Treasure> treasureDetal, Dictionary<int, int> usersTreasure)
    {

        // Create a table to store data for the DropDownList control.
        DataTable dt = new DataTable();

        // Define the columns of the table.
        dt.Columns.Add(new DataColumn("TextField", typeof(String)));
        dt.Columns.Add(new DataColumn("ValueField", typeof(int)));
        Dictionary<int, TreasureHunt.Treasure>.KeyCollection treasureKey = treasureDetal.Keys;
        dt.Rows.Add(CreateRow("", 0, dt));
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;
            if (usersTreasure.ContainsKey(treasure.TreasureId) && usersTreasure[treasure.TreasureId] >= 1)
            {
                dt.Rows.Add(CreateRow(treasure.TreasureName, treasure.TreasureId, dt));
            }
           
        }
        // Populate the table with sample values.

        // Create a DataView from the DataTable to act as the data source
        // for the DropDownList control.
        DataView dv = new DataView(dt);
        return dv;

    }

    DataRow CreateRow(String Text, int Value, DataTable dt)
    {

        // Create a DataRow using the DataTable defined in the 
        // CreateDataSource method.
        DataRow dr = dt.NewRow();

        // This DataRow contains the ColorTextField and ColorValueField 
        // fields, as defined in the CreateDataSource method. Set the 
        // fields with the appropriate value. Remember that column 0 
        // is defined as ColorTextField, and column 1 is defined as 
        // ColorValueField.
        dr[0] = Text;
        dr[1] = Value;

        return dr;

    }

    private class User_Treasure
    {
        int m_TreasureId;
        string m_TreasureName;
        int m_Piece;
        public int TreasureId
        {
            get { return m_TreasureId; }
            set { m_TreasureId = value; }
        }
        public string TreasureName
        {
            get { return m_TreasureName; }
            set { m_TreasureName = value; }
        }
        public int Piece
        {
            get { return m_Piece; }
            set { m_Piece = value; }
        }
    }
}
