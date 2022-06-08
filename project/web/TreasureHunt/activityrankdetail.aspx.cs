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
using System.Web.Configuration;
using System.Text;
using System.Collections.Generic;

public partial class TreasureHunt_activityrankdetail : System.Web.UI.Page
{
    private TreasureHunt treasureHunt;
    private string kmwebsysSite = ConfigurationManager.AppSettings["myURL"].ToString();
    protected string avtivityid = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        SetPageDetal();
    }

    private void SetPageDetal()
    {
        treasureHunt = new TreasureHunt("");
        int pageSize =15;
        int pageNumber = 1;
        int avtivityId = 0;
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 1 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", "0"));
        avtivityid = avtivityId.ToString();
        treasureHunt.SetActivity(avtivityId);
        if (avtivityId == 3)
        {
            HyperLink1.Visible = true;
        }
        if (!IsPostBack)
        {
            pageSize = (WebUtility.GetStringParameter("PageSize", string.Empty) == "") ? 15 : Convert.ToInt32(WebUtility.GetStringParameter("PageSize", string.Empty));
            pageNumber = (WebUtility.GetStringParameter("pagenumber", string.Empty) == "") ? 1 : Convert.ToInt32(WebUtility.GetStringParameter("pagenumber", string.Empty));
            TreasureCategory.DataSource = CreateDataSource();
            TreasureCategory.DataTextField = "TextField";
            TreasureCategory.DataValueField = "ValueField";
            TreasureCategory.DataBind();
            TreasureCategory.SelectedIndex = 0;
        }
        else
        {
            pageSize =  Convert.ToInt32(PageSizeDDL.SelectedValue);
            pageNumber = Convert.ToInt32(PageNumberDDL.SelectedValue);
        }

        Dictionary<int, TreasureHunt.Treasure> treasureDetal = treasureHunt.GetTreasureDetal();

        SetAllFindedTreasure(avtivityId, treasureDetal);
        string startTime = "";
        string endTime = "";
        DateTime timeTemp = new DateTime();
        if (TxtStartDate.Text != "" && TxtEndDate.Text != "")
        {
            if (DateTime.TryParseExact(TxtStartDate.Text, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp) && DateTime.TryParseExact(TxtEndDate.Text, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp))
            {
                startTime = TxtStartDate.Text;
                endTime = TxtEndDate.Text;
            }
        }
        string isForQuery = "";
        if (Request.Form.GetValues("FroQuery") != null)
            isForQuery = (Request.Form.GetValues("FroQuery")[0].ToString());
        if (!string.IsNullOrEmpty(isForQuery))
        {
            if (isForQuery.CompareTo("true") == 0)
                pageNumber = 1;
        }
        double probability = -1;
        if(!string.IsNullOrEmpty(TextBoxProbability.Text))
        {
            probability = Convert.ToDouble(TextBoxProbability.Text)/100;
        }
        string queryMember = "";
        if (!string.IsNullOrEmpty(TextBoxMember.Text))
            queryMember = TextBoxMember.Text;
        int treasureId = 0;
        treasureId = Convert.ToInt32(TreasureCategory.SelectedValue);
        int totalsuits = -1;
        if (!string.IsNullOrEmpty(TextBoxSuits.Text))
            totalsuits = Convert.ToInt32(TextBoxSuits.Text);
        int totalsuitUpper = -1;
        if(!string.IsNullOrEmpty(TextBoxSuitsUpper.Text))
            totalsuitUpper = Convert.ToInt32(TextBoxSuitsUpper.Text);
        int totalsetsLowerBound = 0;
        if (!string.IsNullOrEmpty(TextTreasureSetsLowerBound.Text))
            totalsetsLowerBound = Convert.ToInt32(TextTreasureSetsLowerBound.Text);
        int totalsetsUpperBound = 999;
        if (!string.IsNullOrEmpty(TextTreasureSetsUpperBound.Text))
            totalsetsUpperBound = Convert.ToInt32(TextTreasureSetsUpperBound.Text);
        if (totalsetsUpperBound < 0 || totalsetsLowerBound < 0)
            Response.Write("<script>alert(\"獲得寶物區間錯誤,請重新輸入\");history.go(-1);</script>");
        if (totalsetsLowerBound > totalsetsUpperBound)
        {
            int tempUpperBound = totalsetsUpperBound;
            totalsetsLowerBound = totalsetsUpperBound;
            totalsetsUpperBound = tempUpperBound;
        }
        bool reverse = false;
        if (ReverseChecked.Checked)
            reverse = true;
        bool onlyGetBox=false;
        string getBoxString = "";
        if (OnlyGetBox.Checked)
        {
            onlyGetBox = true;
            getBoxString = "&states=getbox";
        }
        string urlTemp = kmwebsysSite + "/treasureHunt/activityrankdetalexport.aspx?avtivityid=" + avtivityId.ToString() + "&probability=" + probability + "&querymember=" + HttpUtility.UrlEncode(queryMember) + "&treasureid=" + treasureId;
        urlTemp += "&totalsuits=" + totalsuits + "&totalsuitupper=" + totalsuitUpper + "&totalsetslowerbound=" + totalsetsLowerBound + "&totalsetsupperbound=" + totalsetsUpperBound ;
        urlTemp += "&starttime=" + startTime + "&endtime=" + endTime + "&reverse=" + reverse.ToString() + getBoxString;
        linkExport.NavigateUrl = urlTemp;
        IList topList = treasureHunt.GetTopAndQuery(pageNumber, pageSize, 0, queryMember, treasureId, totalsuits, totalsuitUpper, probability, totalsetsLowerBound, totalsetsUpperBound, startTime, endTime, reverse, onlyGetBox); 
        int pageCount = 0;
        int total = treasureHunt.GetTotalGameMunberTopImpl(0, queryMember, treasureId, totalsuits, totalsuitUpper, probability, totalsetsLowerBound, totalsetsUpperBound, startTime, endTime, onlyGetBox);
        if (topList.Count > 0)
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
                PreviousLink.Text = "<a href=\"javascript:PrePage();\" >上一頁 &nbsp;</a>";
            }
            else
            {
                PreviousLink.Text = "上一頁  &nbsp;";
                PreviousText.Enabled = false;
            }
            if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
            {
                NextLink.Text = "<a href=\"javascript:NextPage();\" >下一頁 &nbsp;</a>";
                //NextLink.NavigateUrl = "activityrankdetail.aspx?PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString()+"&avtivityid=" + avtivityId.ToString();
            }
            else
            {
                NextLink.Text = "下一頁 &nbsp;";
                NextText.Enabled = false;
            }
            string link = "";
            StringBuilder sb = new StringBuilder();

            sb.Append("<table class=\"type02\" summary=\"結果條列式\">");
            sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
            sb.Append("<th scope=\"col\" width=\"10%\">帳號</th>");
            sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
            sb.Append("<th scope=\"col\" width=\"5%\">收集<br/>套數</th>");
            if (avtivityId == 2)
            {
                sb.Append("<th scope=\"col\" width=\"10%\">尚未<br/>兌換</th>");
                sb.Append("<th scope=\"col\" width=\"10%\">已兌換</th>");
            }
            else
            {
                sb.Append("<th scope=\"col\" width=\"5%\">抽獎<br/>次數</th>");
            }
            sb.Append("<th scope=\"col\" width=\"8%\">有效閱讀</th>");
            sb.Append("<th scope=\"col\" width=\"8%\">獲得數</th>");
            sb.Append("<th scope=\"col\" width=\"15%\">獲得種類</th>");
            if (avtivityId == 3)
            {
                sb.Append("<th scope=\"col\" width=\"5%\">交換數</th>");
                sb.Append("<th scope=\"col\" width=\"8%\">實際寶物</th>");
                sb.Append("<th scope=\"col\" width=\"15%\">實際擁有種類</th>");
            }
            
            sb.Append("</tr>");


            for (int i = 0; i < topList.Count; i++)
            {
                sb.Append("<tr><td align=\"center\">" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                TreasureHunt.Treasure_Top topObj = new TreasureHunt.Treasure_Top();
                topObj = (TreasureHunt.Treasure_Top)topList[i];
                sb.Append("<td align=\"center\">" + "<a href=\"" + kmwebsysSite + "/treasureHunt/treasurelog.aspx?type=all&accountid=" + topObj.AccountId.ToString() + "&avtivityid=" + avtivityId + getBoxString + "\" >" + topObj.LoginId + "</a></td>");
                if (!string.IsNullOrEmpty(topObj.NickName) && !string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + topObj.RealName+"("+topObj.NickName + ")" + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.NickName))
                {
                    sb.Append("<td align=\"center\">" + "(" + topObj.NickName + ")" + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + topObj.RealName  + "</td>");
                }
                else
                {
                    sb.Append("<td>&nbsp;</td>");
                }
                sb.Append("<td align=\"center\">" + "<a href=\"" + kmwebsysSite + "/treasureHunt/treasurelog.aspx?type=get&accountid=" + topObj.AccountId.ToString() + "&avtivityid=" + avtivityId + "\" >" + topObj.TotalSuit + "</a></td>");
                int voteForLottert = treasureHunt.GetUserVotesLotteryCount(avtivityId, topObj.AccountId);
                if (avtivityId == 2)
                {
                    sb.Append("<td align=\"center\">" + (topObj.TotalSuit - voteForLottert).ToString() + "</td>");
                    sb.Append("<td align=\"center\">" + "<a href=\"" + kmwebsysSite + "/treasureHunt/forLottery.aspx?type=get&accountid=" + topObj.AccountId.ToString() + "&avtivityid=" + avtivityId + "\" >" + voteForLottert.ToString() + "</a></td>");
                }
                else
                {
                    sb.Append("<td align=\"center\">" + topObj.Lottery.ToString() + "</td>");
                }
                    sb.Append("<td align=\"center\">" + topObj.TotalReadCount + "</td>");
                sb.Append("<td align=\"center\">" + topObj.Total_Treasure.ToString() + "</td>");
                string ss = "<td align=\"center\">";
                //取得各user的寶物細項
                Dictionary<int, int> treasureCount = treasureHunt.GetUsetAllKindTreasureCount(avtivityId, topObj.AccountId, startTime,endTime);
                foreach (KeyValuePair<int, TreasureHunt.Treasure> userItem in treasureDetal)
                {
                    TreasureHunt.Treasure treasure = userItem.Value;
                    if (treasure.TreasureName.Length > 3)
                    {
                        ss += treasure.TreasureName.Substring(3, 1) + ":";
                    }
                    else
                    {
                        ss += treasure.TreasureName + ":";
                    }
                    if (treasureCount.ContainsKey(userItem.Key))
                    {
                        ss += treasureCount[userItem.Key].ToString();
                    }
                    else
                    {
                        ss += "0";
                    }
                    ss += "&nbsp;";
                }
                ss += "</td>";
                sb.Append(ss);
                if (avtivityId == 3)
                {
                    sb.Append("<td align=\"center\">" + topObj.Change_Treasure.ToString() + "</td>");
                    sb.Append("<td align=\"center\">" + topObj.Real_Treasure.ToString() + "</td>");
                
                    Dictionary<int, int> usersTreasure = treasureHunt.GetUsersTreasure(topObj.AccountId);
                    ss = "<td align=\"center\">";
                    foreach (KeyValuePair<int, TreasureHunt.Treasure> userItem in treasureDetal)
                    {
                        TreasureHunt.Treasure treasure = userItem.Value;
                        if (treasure.TreasureName.Length > 3)
                        {
                            ss += treasure.TreasureName.Substring(3, 1) + ":";
                        }
                        else
                        {
                            ss += treasure.TreasureName + ":";
                        }
                        if (usersTreasure.ContainsKey(userItem.Key))
                        {
                            ss += usersTreasure[userItem.Key].ToString();
                        }
                        else
                        {
                            ss += "0";
                        }
                        ss += "&nbsp;";
                    }
                    ss += "</td>";
                    sb.Append(ss);
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
            TableText.Text = sb.ToString();
        }
        else
        {
            StringBuilder sb = new StringBuilder();

            PageNumberText.Text = "0";
            TotalPageText.Text = "0";
            TotalRecordText.Text = "0";

            sb.Append("<table class=\"type02\" summary=\"結果條列式\">");
            sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
            sb.Append("<th scope=\"col\" width=\"10%\">帳號</th>");
            sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
            sb.Append("<th scope=\"col\" width=\"10%\">已收集套數</th>");
            if (avtivityId == 2)
            {
                sb.Append("<th scope=\"col\" width=\"10%\">尚未兌換套數</th>");
                sb.Append("<th scope=\"col\" width=\"10%\">已兌換套數</th>");
            }
            else
            {
                sb.Append("<th scope=\"col\" width=\"10%\">抽獎次數</th>");
            }
            sb.Append("<th scope=\"col\" width=\"10%\">有效閱讀數</th>");
            if (avtivityId == 3)
            {
                sb.Append("<th scope=\"col\" width=\"10%\">獲得寶物數</th>");
                sb.Append("<th scope=\"col\" width=\"20%\">獲得種類&數量</th>");
            }
            sb.Append("</tr>");
            sb.Append("</table>");

            TableText.Text = sb.ToString();
        }

        
    }

    private void SetAllFindedTreasure(int actId, Dictionary<int, TreasureHunt.Treasure> treasureDetal)
    {
        Dictionary<int,int> treasureCount = treasureHunt.GetAllFindedTreasureCount(actId);
        string ss = "&nbsp;&nbsp;";
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;
            ss += treasure.TreasureName + ":";
            if (treasureCount.ContainsKey(item.Key))
            {
                ss += treasureCount[item.Key].ToString();
            }
            else
            {
                ss += "0";
            }
            ss += "&nbsp;&nbsp;&nbsp;";
        }
        AllFindedTreasure.Text = ss;
        treasureCount.Clear();
        ss = "&nbsp;&nbsp;";
        treasureCount = treasureHunt.GetAllUserBackageTreasureCount(actId);
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;
            ss += treasure.TreasureName + ":";
            if (treasureCount.ContainsKey(item.Key))
            {
                ss += treasureCount[item.Key].ToString();
            }
            else
            {
                ss += "0";
            }
            ss += "&nbsp;&nbsp;&nbsp;";
        }
        AllUserTreasure.Text = ss;
    }

    ICollection CreateDataSource()
    {

        // Create a table to store data for the DropDownList control.
        DataTable dt = new DataTable();

        // Define the columns of the table.
        dt.Columns.Add(new DataColumn("TextField", typeof(String)));
        dt.Columns.Add(new DataColumn("ValueField", typeof(int)));
        Dictionary<int, TreasureHunt.Treasure> treasureDetal = treasureHunt.GetTreasureDetal();
        Dictionary<int, TreasureHunt.Treasure>.KeyCollection treasureKey = treasureDetal.Keys;
        dt.Rows.Add(CreateRow("全部", 0, dt));
        foreach (KeyValuePair<int, TreasureHunt.Treasure> item in treasureDetal)
        {
            TreasureHunt.Treasure treasure = item.Value;
            dt.Rows.Add(CreateRow(treasure.TreasureName, treasure.TreasureId, dt));
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
