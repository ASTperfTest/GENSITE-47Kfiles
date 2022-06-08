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

public partial class TreasureHunt_treasurelog : System.Web.UI.Page
{
    private TreasureHunt treasureHunt;
    int pageSize;
    int pageNumber;
    protected int avtivityId = 0;
    private string kmwebsysSite = ConfigurationManager.AppSettings["myURL"].ToString();
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            pageSize = (WebUtility.GetStringParameter("PageSize", string.Empty) == "") ? 15 : Convert.ToInt32(WebUtility.GetStringParameter("PageSize", string.Empty));
            pageNumber = (WebUtility.GetStringParameter("pagenumber", string.Empty) == "") ? 1 : Convert.ToInt32(WebUtility.GetStringParameter("pagenumber", string.Empty));
        }
        else
        {
            pageSize = Convert.ToInt32(PageSizeDDL.SelectedValue);
            pageNumber = Convert.ToInt32(PageNumberDDL.SelectedValue);
        }
        if (Request.Form.GetValues("ForQuery") != null && !string.IsNullOrEmpty(Request.Form.GetValues("ForQuery")[0].ToString()) && Request.Form.GetValues("ForQuery")[0].ToString()=="true")
            pageNumber = 1;
        string accountId = WebUtility.GetStringParameter("accountid", string.Empty).ToLower();
        string type = WebUtility.GetStringParameter("type", string.Empty).ToLower();
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 0 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", "0"));
        treasureHunt = new TreasureHunt("");
        treasureHunt.SetActivity(avtivityId);
        if (!string.IsNullOrEmpty(accountId) && type=="get")
        {
            SetUserDetal(accountId);
            SetUserTreasure(accountId);
            SetUserGetTreasureLog(accountId, avtivityId, 2);
        }
        else if (!string.IsNullOrEmpty(accountId) && type == "all")
        {
            QueryForGetBox.Visible = true;
            QueryForALL.Visible = true;
            ChangeALL.Visible = true;
            int states = -1;
            if (WebUtility.GetStringParameter("states", string.Empty).ToLower() == "getbox")
                states = 0;
            SetUserDetal(accountId);
            SetUserTreasureLog(accountId, avtivityId, states);
        }
    }

    private void SetUserDetal(string accountId)
    {
        int id = Convert.ToInt32(accountId);
        Dictionary<string, string> userData = treasureHunt.GetUserData(id);
        if(userData["loginId"]!=null)
            loginid.Text = userData["loginId"];
        if (userData["name"] != null)
            userName.Text = userData["name"];
        if (userData["mail"] != null)
            userMail.Text = userData["mail"];
    }

    private void SetUserTreasure(string accountId)
    {
        int id = Convert.ToInt32(accountId);
        string ss = "";
        Dictionary<int, TreasureHunt.Treasure> treasureDetal = treasureHunt.GetTreasureDetal();
        Dictionary<int, int> usersTreasure = treasureHunt.GetUsersTreasure(id);
        int totalSuit = 0;
        int maxSiut = 5;
        ss = "<table>";
        Dictionary<int, TreasureHunt.Treasure>.KeyCollection treasureKey = treasureDetal.Keys;
        ss += "<tr>";
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;

            ss += "<td align=\"center\" width=\"96\" >" + treasure.TreasureName + "</td> ";
        }
        ss += "</tr>";
        ss += "<tr>";
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;

            ss += "<td align=\"center\"><img src=\"" + kmwebsysSite + "/TreasureHunt/image/" + treasure.Icon + "\" ></td> ";
        }
        ss += "</tr>";
        bool flag = false;
        ss += "<tr>";
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;

            ss += "<td align=\"center\">數量:";
            if (usersTreasure.ContainsKey(item.Key))
            {
                if (flag)
                {
                    totalSuit = Math.Min(totalSuit, usersTreasure[item.Key]);
                }
                else
                {
                    totalSuit = usersTreasure[item.Key];
                    flag = true;
                }

                ss += (usersTreasure[item.Key]).ToString();
            }
            else
            {
                totalSuit = 0;
                flag = true;
                ss += "0";
            }

            ss += "</td> ";
        }
        ss += "</tr>";
        ss += "</table>";

        suit.Text = "已收集套數:" + totalSuit.ToString();
        treasureTable.Text = ss;
    }

    private void SetUserGetTreasureLog(string accountId, int activityId, int states)
    {
        int id = Convert.ToInt32(accountId);
        string ss = "";
        int totalSuit = 0;
        int maxSiut = 5;
        ss = "<table width=\"60%\"><tr><th  width=\"20%\">來源</th><th  width=\"40%\">標題</th><th  width=\"20%\">取得寶物</th><th  width=\"20%\">取得時間</th></tr>";
        int pageCount = 0;
        IList treasureLogList = treasureHunt.GetSearchTreasureRecord(id, activityId, states, pageNumber, pageSize);
        foreach (TreasureHunt.TreasureLogDetal treasureLogDetal in treasureLogList)
        {
            ss += "<tr>";
            if (!string.IsNullOrEmpty(treasureLogDetal.UnitName))
            {
                ss += "<td align=\"center\">" + treasureLogDetal.UnitName + "</td>";
            }
            else if (treasureLogDetal.States == 3)
            {
                ss += "<td align=\"center\">使用於交換</td>";
            }
            else if (treasureLogDetal.States == 4)
            {
                ss += "<td align=\"center\">交換獲得</td>";
            }
            else
            {
                ss += "<td>&nbsp;</td>";
            }
            ss += "<td><a href=\"#\" onclick=\"window.open('" + kmwebsysSite +treasureLogDetal.Url +  "')\" > " + treasureLogDetal.DocumentTitle + "</a></td>";
            ss += "<td align=\"center\">" + treasureLogDetal.TreasureName + "</td>";
            ss += "<td align=\"center\">" + treasureLogDetal.GetDate + "</td>";
            ss += "</tr>";
        }
        ss += "</table>";
        TableText.Text = ss;

        int total = treasureHunt.getTreasureHuntLogCount(id, activityId, states);
        pageCount = Convert.ToInt32((total / pageSize + 0.999));
        if (treasureLogList.Count > 0)
        {
            pageCount = Convert.ToInt32((total / pageSize + 0.999));
            if (pageCount < pageNumber)
            {
                pageNumber = pageCount;
            }
            PageNumberText.Text = pageNumber.ToString();
            TotalPageText.Text = pageCount.ToString();
            TotalRecordText.Text = total.ToString();
            if (pageSize == 15)
            {
                PageSizeDDL.SelectedIndex = 0;
            }
            else if (pageSize == 30)
            {
                PageSizeDDL.SelectedIndex = 1;
            }
            else if (pageSize == 50)
            {
                PageSizeDDL.SelectedIndex = 2;
            }
            ListItem item = default(ListItem);
            int j = 0;
            PageNumberDDL.Items.Clear();
            for (j = 0; j <= pageCount - 1; j++)
            {
                item = new ListItem();
                item.Value = (j + 1).ToString();
                item.Text = (j + 1).ToString();
                if (pageNumber == (j + 1))
                {
                    item.Selected = true;
                }
                PageNumberDDL.Items.Insert(j, item);
                item = null;
            }


            if (pageNumber > 1)
            {
                PreviousLink.NavigateUrl = "treasurelog.aspx?type=get&accountid=" + accountId + "&PageNumber=" + (pageNumber - 1).ToString() + "&PageSize=" + pageSize.ToString() + "&avtivityid=" + avtivityId;
            }
            else
            {
                PreviousText.Enabled = false;
                PreviousLink.NavigateUrl = "";
            }
            if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
            {
                NextLink.NavigateUrl = "treasurelog.aspx?type=get&accountid=" + accountId + "&PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString() + "&avtivityid=" + avtivityId;
            }
            else
            {
                NextText.Enabled = false;
                NextLink.NavigateUrl = "";
            }

        }
    }


    private void SetUserTreasureLog(string accountId, int activityId, int states)
    {

        int id = Convert.ToInt32(accountId);
        string ss = "";
        int totalSuit = 0;
        int maxSiut = 5;
        ss = "<table  width=\"60%\"><tr><th width=\"20%\" align=\"center\">來源</th><th width=\"35%\" align=\"center\">文章標題&交換</th>";
        ss += "<th width=\"15%\">是否出現寶物</th><th width=\"15%\" align=\"center\" >寶物名稱</th><th width=\"15%\" align=\"center\">閱覽時間</th></tr>";
        int pageCount = 0;
        string stateTemp="";
        if (Request.Form.GetValues("ReverseQuery") != null &&  !string.IsNullOrEmpty(Request.Form.GetValues("ReverseQuery")[0].ToString()))
        {
            stateTemp = Request.Form.GetValues("ReverseQuery")[0].ToString();
            if (stateTemp.CompareTo("true") == 0)
            {
                states = 0;
            }
        }
        if (Request.Form.GetValues("ForChange") != null && !string.IsNullOrEmpty(Request.Form.GetValues("ForChange")[0].ToString()))
        {
            stateTemp = Request.Form.GetValues("ForChange")[0].ToString();
            if (stateTemp.CompareTo("true") == 0)
            {
                states = 4;
            }
        }
        string stateUrl="";
        if (states == 0)
        {
            stateUrl = "&states=getbox";
        }
        else if (states ==4)
        {
            stateUrl = "&states=change";
        }
        IList treasureLogList = treasureHunt.GetSearchTreasureRecord(id, activityId, states, pageNumber, pageSize);
        int total = treasureHunt.getTreasureHuntLogCount(id, activityId, states);
        int denominator = treasureHunt.getTreasureHuntLogCount(id, activityId, -2);
        int getCount = treasureHunt.getTreasureHuntLogCount(id, activityId, 2);
        foreach (TreasureHunt.TreasureLogDetal treasureLogDetal in treasureLogList)
        {
            ss += "<tr>";
            if (!string.IsNullOrEmpty(treasureLogDetal.UnitName))
            {
                ss += "<td align=\"center\">" + treasureLogDetal.UnitName + "</td>";
            }
            else
            {
                ss += "<td>&nbsp;</td>";
            }
            switch (treasureLogDetal.States)
            {
                case 3:
                    ss += "<td>使用寶物交換</td>";
                    break;
                case 4:
                    ss += "<td>交換後獲得的寶物</td>";
                    break;
                default:
                    ss += "<td><a href=\"#\" onclick=\"window.open('" + kmwebsysSite + treasureLogDetal.Url + "')\" > " + treasureLogDetal.DocumentTitle + "</a></td>";
                    break;
            }
            
            string temp = "";
            switch (treasureLogDetal.States)
            {
                case 1:
                    temp = "否";
                    break;
                case 0:
                    temp = "是(未領取)";
                    break;
                case 2:
                    temp = "是";
                    break;
                case 3:
                    temp = "否(交換)";
                    break;
                case 4:
                    temp = "是(交換)";
                    break;
                default:
                    temp = "否";
                    break;
            }
            ss += "<td align=\"center\">" + temp + "</td>";
            if (!string.IsNullOrEmpty(treasureLogDetal.TreasureName))
            {
                ss += "<td align=\"center\">" + treasureLogDetal.TreasureName + "</td>";
            }
            else
            {
                ss += "<td>&nbsp;</td>";
            }
            ss += "<td align=\"center\">" + treasureLogDetal.GetDate + "</td>";
            ss += "</tr>";
        }
        ss += "</table>";
        TableText.Text = ss;
        if (states == 0)
        {
            suit.Text = "未領取次數:" + total;
        }
        else if (states == 4)
        {
            suit.Text = "交換次數";
        }else
        {
            suit.Text = "出現寶物機率:" + getCount.ToString() + "/" + denominator;
        }


        
        pageCount = Convert.ToInt32((total / pageSize + 0.999));
        if ((total % pageSize) ==0)
            pageCount = Convert.ToInt32((total / pageSize));
        if (treasureLogList.Count > 0)
        {
            pageCount = Convert.ToInt32((total / pageSize + 0.999));
            if ((total % pageSize) == 0)
                pageCount = Convert.ToInt32((total / pageSize));
            if (pageCount < pageNumber)
            {
                pageNumber = pageCount;
            }
            PageNumberText.Text = pageNumber.ToString();
            TotalPageText.Text = pageCount.ToString();
            TotalRecordText.Text = total.ToString();
            if (pageSize == 15)
            {
                PageSizeDDL.SelectedIndex = 0;
            }
            else if (pageSize == 30)
            {
                PageSizeDDL.SelectedIndex = 1;
            }
            else if (pageSize == 50)
            {
                PageSizeDDL.SelectedIndex = 2;
            }
            ListItem item = default(ListItem);
            int j = 0;
            PageNumberDDL.Items.Clear();
            for (j = 0; j <= pageCount - 1; j++)
            {
                item = new ListItem();
                item.Value = (j + 1).ToString();
                item.Text = (j + 1).ToString();
                if (pageNumber == (j + 1))
                {
                    item.Selected = true;
                }
                PageNumberDDL.Items.Insert(j, item);
                item = null;
            }


            if (pageNumber > 1)
            {
                PreviousLink.NavigateUrl = "treasurelog.aspx?type=all&accountid=" + accountId + "&PageNumber=" + (pageNumber - 1).ToString() + "&PageSize=" + pageSize.ToString() + "&avtivityid=" + avtivityId + stateUrl;
            }
            else
            {
                PreviousText.Enabled = false;
                PreviousLink.NavigateUrl = "";
            }
            if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
            {
                NextLink.NavigateUrl = "treasurelog.aspx?type=all&accountid=" + accountId + "&PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString() + "&avtivityid=" + avtivityId + stateUrl;
            }
            else
            {
                NextText.Enabled = false;
                NextLink.NavigateUrl = "";
            }

        }
    }
}
