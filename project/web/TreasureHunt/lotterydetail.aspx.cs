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
using System.Text;

public partial class TreasureHunt_lotterydetail : System.Web.UI.Page
{
    private TreasureHunt treasureHunt;
    private string kmwebsysSite = ConfigurationManager.AppSettings["myURL"].ToString();
    protected string activityId = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        SetPageDetal();
    }

    private void SetPageDetal()
    {
        treasureHunt = new TreasureHunt("");
        int pageSize = 15;
        int pageNumber = 1;
        int avtivityId = 0;
        string giftSortId = "";
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 1 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", "0"));
        activityId = avtivityId.ToString();
        treasureHunt.SetActivity(avtivityId);

        IList GiftList = treasureHunt.GetAllGiftVote(Convert.ToInt32(activityId));
        if (!IsPostBack)
        {
            pageSize = (WebUtility.GetStringParameter("PageSize", string.Empty) == "") ? 15 : Convert.ToInt32(WebUtility.GetStringParameter("PageSize", string.Empty));
            pageNumber = (WebUtility.GetStringParameter("pagenumber", string.Empty) == "") ? 1 : Convert.ToInt32(WebUtility.GetStringParameter("pagenumber", string.Empty));
            GiftCategory.DataSource = CreateDataSource(GiftList);
            GiftCategory.DataTextField = "TextField";
            GiftCategory.DataValueField = "ValueField";
            GiftCategory.DataBind();
            GiftCategory.SelectedIndex = 0;
        }
        else
        {
            pageSize = Convert.ToInt32(PageSizeDDL.SelectedValue);
            pageNumber = Convert.ToInt32(PageNumberDDL.SelectedValue);
        }
        giftSortId = GiftCategory.SelectedValue;
        LabelTableName.Text = "<font size=\"4px\">目前列表:" + GiftCategory.SelectedItem + "投票清單</font>";
        string ss = "";
        ss += "<table> ";
        int index = 0;
        ss += "<tr>";
        foreach (TreasureHunt.LotteryGift temp in GiftList)
        {
            if (giftSortId == "")
                giftSortId = temp.GiftId;
            if ((index % 6) == 0 && index != 0)
                ss += "<tr>";
            ss += "<td width=\"130px\" align=\"center\">";
            ss += "<span class=\"present_title\">" + temp.GiftName + " </span><br/>";
            ss += "共<span class=\"chance\" >" + temp.Votes + "</span>次 </td> ";
            index++;
            if ((index % 6) == 0 && GiftList.Count > index)
                ss += "</tr>";
        }
        ss += "</tr>";
        ss += "</table>";
        LabelGiftVotes.Text = ss;

        SetLotteryTop(avtivityId, giftSortId, pageNumber, pageSize);

    }
    private void SetLotteryTop(int avtivityId,string giftId,int pageNumber,int pageSize)
    {
        string userName = TextBoxMember.Text;
        IList lotteryTopList = treasureHunt.GetGiftVoteUsers(avtivityId, giftId,userName, pageNumber, pageSize);
        int total = treasureHunt.GetVoteCountById(avtivityId, giftId);
        int pageCount = 0;
        string urlTemp = kmwebsysSite + "/treasureHunt/lotterydetailexport.aspx?avtivityid=" + avtivityId.ToString() + "&giftid=" + giftId + "&querymember=" + HttpUtility.UrlEncode(userName);
        linkExport.NavigateUrl = urlTemp;
        if (lotteryTopList.Count > 0)
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
            sb.Append("<th scope=\"col\" width=\"10%\">換取次數</th>");
            sb.Append("</tr>");


            for (int i = 0; i < lotteryTopList.Count; i++)
            {
                sb.Append("<tr><td align=\"center\">" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                TreasureHunt.Treasure_Top topObj = new TreasureHunt.Treasure_Top();
                topObj = (TreasureHunt.Treasure_Top)lotteryTopList[i];
                sb.Append("<td align=\"center\">" + "<a href=\"" + kmwebsysSite + "/treasureHunt/treasurelog.aspx?type=all&accountid=" + topObj.AccountId.ToString() + "&avtivityid=" + avtivityId + "\" >" + topObj.LoginId + "</a></td>");
                if (!string.IsNullOrEmpty(topObj.NickName) && !string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + topObj.RealName + "(" + topObj.NickName + ")" + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.NickName))
                {
                    sb.Append("<td align=\"center\">" + "(" + topObj.NickName + ")" + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + topObj.RealName + "</td>");
                }
                else
                {
                    sb.Append("<td>&nbsp;</td>");
                }
                sb.Append("<td align=\"center\">" + "<a href=\"" + kmwebsysSite + "/treasureHunt/forLottery.aspx?type=get&accountid=" + topObj.AccountId.ToString() + "&avtivityid=" + avtivityId + "\" >" + topObj.Lottery + "</a></td>");

                sb.Append("</tr>");

            }
            sb.Append("</table>");
            TableText.Text = sb.ToString();
        }
        else
        {
            PageNumberText.Text = "0";
            TotalPageText.Text = "0";
            TotalRecordText.Text = "0";
            StringBuilder sb = new StringBuilder();
            sb.Append("<table class=\"type02\" summary=\"結果條列式\">");
            sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
            sb.Append("<th scope=\"col\" width=\"10%\">帳號</th>");
            sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
            sb.Append("<th scope=\"col\" width=\"10%\">換取次數</th>");
            sb.Append("</tr>");
            sb.Append("</table>");
            TableText.Text = sb.ToString();
        }
    }

    ICollection CreateDataSource(IList list)
    {

        // Create a table to store data for the DropDownList control.
        DataTable dt = new DataTable();

        // Define the columns of the table.
        dt.Columns.Add(new DataColumn("TextField", typeof(String)));
        dt.Columns.Add(new DataColumn("ValueField", typeof(String)));
        foreach (TreasureHunt.LotteryGift temp in list)
        {
            dt.Rows.Add(CreateRow(temp.GiftName, temp.GiftId, dt));
        }
        // Populate the table with sample values.

        // Create a DataView from the DataTable to act as the data source
        // for the DropDownList control.
        DataView dv = new DataView(dt);
        return dv;

    }

    DataRow CreateRow(String Text, String Value, DataTable dt)
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
