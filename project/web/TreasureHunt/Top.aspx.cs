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

public partial class TreasureHunt_Top : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        SetPageDetal();
    }

    private void SetPageDetal()
    {
        int pageSize =15;
        int pageNumber = 1;
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
        TreasureHunt treasureHunt = new TreasureHunt("");
        int avtivityId = 2;
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 2 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", "0"));
        treasureHunt.SetActivity(avtivityId);
        IList topList = treasureHunt.GetTop(pageNumber, pageSize, avtivityId, "", -1, -1,-1, -1,-1,-1);
        int pageCount = 0;
        int total = treasureHunt.GetTotalGameMunber(-1);
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
                PreviousLink.NavigateUrl = "Top.aspx?PageNumber=" + (pageNumber - 1).ToString() + "&PageSize=" + pageSize.ToString() + "&avtivityid=" + avtivityId.ToString();
            }
            else
            {
                PreviousText.Enabled = false;
                PreviousLink.NavigateUrl = "";
            }
            if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
            {
                NextLink.NavigateUrl = "Top.aspx?PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString() + "&avtivityid=" + avtivityId.ToString();
            }
            else
            {
                NextText.Enabled = false;
                NextLink.NavigateUrl = "";
            }
            string link = "";
            StringBuilder sb = new StringBuilder();

            sb.Append("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" id=\"qa\">");
            sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
            sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
            sb.Append("<th scope=\"col\" width=\"15%\">已收集套數</th>");
            if (avtivityId != 1)
            {
                sb.Append("<th scope=\"col\" width=\"15%\">已兌換套數</th>");
            }
            else
            {
                sb.Append("<th scope=\"col\" width=\"15%\">抽獎次數</th>");
            }
            sb.Append("<th scope=\"col\" width=\"20%\">獲得寶物數</th>");
            //sb.Append("</tr>");
            //sb.Append("<tr><td width=\"30\" height=\"30\" style=\"background:url(image/qa_br_tl.gif) no-repeat bottom\">&nbsp;</td>");
            //sb.Append("<td align=\"center\" style=\"background:url(image/qa_br_t.gif) repeat-x bottom\"><img src=\"image/qa_h1.gif\" width=\"423\" height=\"167\"/></td>");
            //sb.Append("<td width=\"30\" style=\"background:url(image/qa_br_tr.gif) no-repeat bottom\">&nbsp;</td>");
            //sb.Append("</tr>");
            

            for (int i = 0; i < topList.Count; i++)
            {
                sb.Append("<tr><td>" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                TreasureHunt.Treasure_Top topObj = new TreasureHunt.Treasure_Top();
                topObj = (TreasureHunt.Treasure_Top)topList[i];
                if (!string.IsNullOrEmpty(topObj.NickName))
                {
                    sb.Append("<td>" + topObj.NickName + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td>" + (topObj.RealName.Substring(0, 1) + "＊" + topObj.RealName.Substring(topObj.RealName.Trim().Length - 1, 1)) + "</td>");
                }
                else
                {
                    sb.Append("<td>&nbsp;</td>");
                }
                sb.Append("<td align=\"center\">" + topObj.TotalSuit + "</td>");
                int voteForLottert = treasureHunt.GetUserVotesLotteryCount(avtivityId, topObj.AccountId);
                if (avtivityId != 1)
                {
                    sb.Append("<td align=\"center\">" +  voteForLottert.ToString() + "</td>");
                }
                else
                {
                    sb.Append("<td align=\"center\">" + topObj.Lottery.ToString() + "</td>");
                }
                sb.Append("<td align=\"center\">" + topObj.Total_Treasure.ToString() + "</td>");
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

            sb.Append("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" id=\"qa\">");
            sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
            sb.Append("<th scope=\"col\" width=\"20%\">姓名｜暱稱</th>");
            sb.Append("<th scope=\"col\" width=\"20%\">已收集套數</th>");
            if (avtivityId != 1)
            {
                sb.Append("<th scope=\"col\" width=\"15%\">已兌換套數</th>");
            }
            else
            {
                sb.Append("<th scope=\"col\" width=\"15%\">抽獎次數</th>");
            }
            sb.Append("<th scope=\"col\" width=\"20%\">獲得寶物數</th>");
            sb.Append("</tr>");
            sb.Append("</table>");

            TableText.Text = sb.ToString();

        }

        
    }
}
