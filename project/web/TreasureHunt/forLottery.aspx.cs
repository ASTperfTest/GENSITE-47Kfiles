using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Collections.Generic;

public partial class TreasureHunt_forLottery : System.Web.UI.Page
{
    private TreasureHunt treasureHunt;
    private string kmwebsysSite = ConfigurationManager.AppSettings["myURL"].ToString();
    protected int avtivityId = 2;
    protected void Page_Load(object sender, EventArgs e)
    {
        string accountId = WebUtility.GetStringParameter("accountid", string.Empty).ToLower();
        string type = WebUtility.GetStringParameter("type", string.Empty).ToLower();
        int avtivityId = 2;
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 0 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", "0"));
        treasureHunt = new TreasureHunt("");
        treasureHunt.SetActivity(avtivityId);
        SetUserDetal(accountId);
        SetUserLoeerty(accountId);
    }

    private void SetUserDetal(string accountId)
    {
        int id = Convert.ToInt32(accountId);
        Dictionary<string, string> userData = treasureHunt.GetUserData(id);
        if (userData["loginId"] != null)
            loginid.Text = userData["loginId"];
        if (userData["name"] != null)
            userName.Text = userData["name"];
        if (userData["mail"] != null)
            userMail.Text = userData["mail"];
    }

    private void SetUserLoeerty(string accountId)
    {
        int id = Convert.ToInt32(accountId);
        string ss = "";
        int totalSuit = 0;
        int forLottery = 0;
        totalSuit = treasureHunt.GetUsersTotalSuit(id);
        IList lotteryList = treasureHunt.GetUserVotesGift(avtivityId, id);
        ss += "<table> ";
        int index = 0;
        ss += "<tr>";
        foreach (TreasureHunt.LotteryGift temp in lotteryList)
        {
            if ((index % 6) == 0 && index != 0)
                ss += "<tr>";
            ss += "<td width=\"130px\" align=\"center\"><img src=\"image/" + temp.Images + "\" width=\"100\" height=\"100\" /><p>";
            ss += "<span class=\"present_title\">" + temp.GiftName + " </span></p>";
            ss += "<p>共<span class=\"chance\" >" + temp .Votes+ "</span>次 </p></td> ";
            forLottery += temp.Votes;
            index++;
            if ((index % 6) == 0 && lotteryList.Count > index)
                ss += "</tr>";
        }
        ss += "</tr>";
        ss += "</table>";
        LabelGiftVotes.Text = ss;
        suit.Text = "已收集套數:" + totalSuit.ToString();
        ForLottery.Text = forLottery.ToString();
    }
}
