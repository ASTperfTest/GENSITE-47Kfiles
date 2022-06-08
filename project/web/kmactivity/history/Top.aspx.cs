using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections;
using System.Text;

public partial class kmactivity_history_Top : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        SetPageDetal();
    }
    private void SetPageDetal()
    {
        int pageSize = 15;
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
        HistoryPicture historyPicture = new HistoryPicture("");
        IList historyTop = historyPicture.GetTop("", "", pageSize, pageNumber);
        StringBuilder sb = new StringBuilder();

        sb.Append("<table width=\"100%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" id=\"rank\">");
        sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">活動得分</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">答題數</th></tr>");
        if (historyTop.Count > 0)
        {
            for (int i = 0; i < historyTop.Count; i++)
            {
                sb.Append("<tr><td>" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                HistoryPicture.HistoryTop topObj = (HistoryPicture.HistoryTop)historyTop[i];
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
                sb.Append("<td align=\"center\">" + topObj.Score + "</td>");
                sb.Append("<td align=\"center\">" + topObj.Count + "</td>");
            }
        }
        sb.Append("</table>");
        TableText.Text = sb.ToString();
        PageNumberText.Text = pageNumber.ToString();
        int totaCount = historyPicture.GetTopCoun("","");
        int pageCount = Convert.ToInt32((totaCount / pageSize + 0.999));
        if ((totaCount % pageSize) == 0)
            pageCount = Convert.ToInt32((totaCount / pageSize));
        TotalPageText.Text = pageCount.ToString();
        ListItem item = default(ListItem);
        PageNumberDDL.Items.Clear();
        for (int j = 0; j <= pageCount - 1; j++)
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
            PreviousLink.NavigateUrl = "top.aspx?a=history&PageNumber=" + (pageNumber - 1).ToString() + "&PageSize=" + pageSize.ToString();
        }
        else
        {
            PreviousText.Enabled = false;
            PreviousLink.NavigateUrl = "";
        }
        if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
        {
            NextLink.NavigateUrl = "top.aspx?a=history&PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString();
        }
        else
        {
            NextText.Enabled = false;
            NextLink.NavigateUrl = "";
        }
        TotalRecordText.Text = totaCount.ToString();
    }
}