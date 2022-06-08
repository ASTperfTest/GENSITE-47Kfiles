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

public partial class TreasureHunt_myspace : System.Web.UI.Page
{
    private ILog log4;
    private TreasureHunt treasureHunt;
    private int avtivityId = 2;
    string connString = System.Web.Configuration.WebConfigurationManager.AppSettings["COAConnectionString"];
    protected string buttonDisable="true";
    protected string voteDisplay = "false";
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
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 2 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", "2"));
        string loginId = Session["memID"].ToString().Trim();
        if (!string.IsNullOrEmpty(Request.Form["__EVENTTARGET"]))
            if (Request.Form["__EVENTTARGET"].Replace("$", "_") == randomChangeTreasure.ClientID)
            {
                RanDomChangeTreasure();
            }
        if (SaveUserData(loginId) == 0)
        {
            Response.Write("<script>alert('連線逾時或尚未登入，請登入會員!');history.go(-1);</script>");
            return;
        }
        treasureHunt = new TreasureHunt(loginId);
        treasureHunt.SetActivity(avtivityId);
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
        IList lotteryList = treasureHunt.GetUserVotesLottery(avtivityId, treasureHunt.account_id);
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
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;
            treasureList.Add(treasure);
        }
        int i = 0;
        while(i<treasureList.Count)
        {
            TreasureHunt.Treasure treasure = new TreasureHunt.Treasure();
            ss += "<tr>";
            if (i < treasureList.Count)
            {
                treasure = (TreasureHunt.Treasure)treasureList[i];
                if (GetUserTrasureCount(usersTreasure, treasure.TreasureId).CompareTo("0") == 0)
                    treasure.Icon = "bean_q.gif";
                ss += "<td width=\"20%\" align=\"center\"><img src=\"image/" + treasure.Icon + "\" /><br />";
                ss += "X" + GetUserTrasureCount(usersTreasure, treasure.TreasureId) + "</td>";
                i = i + 1;
            }
            if (i < treasureList.Count)
            {
                treasure = (TreasureHunt.Treasure)treasureList[i];
                if (GetUserTrasureCount(usersTreasure, treasure.TreasureId).CompareTo("0") == 0)
                    treasure.Icon = "bean_q.gif";
                ss += "<td width=\"20%\" align=\"center\"><img src=\"image/" + treasure.Icon + "\"  /><br />";
                ss += "X" + GetUserTrasureCount(usersTreasure, treasure.TreasureId) + "</td>";
                i = i + 1;
            }
            else
            {
                ss += "<td width=\"20%\" align=\"center\"></td>";
            }
            if (i < treasureList.Count)
            {

                treasure = (TreasureHunt.Treasure)treasureList[i];
                if (GetUserTrasureCount(usersTreasure, treasure.TreasureId).CompareTo("0") == 0)
                    treasure.Icon = "bean_q.gif";
                ss += "<td width=\"20%\" align=\"center\"><img src=\"image/" + treasure.Icon + "\"  /><br />";
                ss += "X" + GetUserTrasureCount(usersTreasure, treasure.TreasureId) + "</td>";
                i = i + 1;
            }
            else
            {
                ss += "<td width=\"20%\" align=\"center\"></td>";
            }
            ss += "<tr>";
        }
        ss += "</table>";
        treasureTable.Text = ss;

        TreasureDropDownList.DataSource = CreateDataSource(treasureDetal, usersTreasure);
        TreasureDropDownList.DataTextField = "TextField";
        TreasureDropDownList.DataValueField = "ValueField";
        TreasureDropDownList.DataBind();
       // TreasureDropDownList.SelectedIndex = 0;
    }

    private string GetUserTrasureCount(Dictionary<int, int> usersTreasure,int id)
    {
        string returnString = "0";
        if (usersTreasure.ContainsKey(id))
        {

            returnString = (usersTreasure[id]).ToString();
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
        IList lotteryList = treasureHunt.GetUserVotesLottery(avtivityId, treasureHunt.account_id);
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
        treasureHunt.SetUserVotesLottery(avtivityId, treasureHunt.account_id, giftId);
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
            Response.End();
        }
        Session["voteLottery"] = "true";
        string sql = @"
               

        ";
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
            if (usersTreasure.ContainsKey(treasure.TreasureId) && usersTreasure[treasure.TreasureId] >= 2)
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
}
