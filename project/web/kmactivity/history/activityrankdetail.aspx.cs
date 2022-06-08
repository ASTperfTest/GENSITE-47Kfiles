using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Collections;
using GSS.Vitals.COA.Data;
using System.Configuration;

public partial class kmactivity_history_activityrankdetail : System.Web.UI.Page
{
    HistoryPicture historyPicture;
    private string kmwebsysSite = ConfigurationManager.AppSettings["myURL"].ToString();
    protected void Page_Load(object sender, EventArgs e)
    {
        historyPicture = new HistoryPicture("");
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        if (WebUtility.GetStringParameter("disableUser", "") != "")
        {
            UpdateDisableUser(1, WebUtility.GetStringParameter("disableUser", ""));
        }
        if (WebUtility.GetStringParameter("unDisableUser", "") != "")
        {
            UpdateDisableUser(0, WebUtility.GetStringParameter("unDisableUser", ""));
        }
        DisplayTable();
    }

    private void DisplayTable()
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
        int scoreUpper = -1;
        if (!string.IsNullOrEmpty(TextScoreUpperBound.Text))
            scoreUpper = Convert.ToInt32(TextScoreUpperBound.Text);
        int scoreLowerBound = 0;
        if (!string.IsNullOrEmpty(TextScoreLowerBound.Text))
            scoreLowerBound = Convert.ToInt32(TextScoreLowerBound.Text);
        int answercountlow = -1;
        if (!string.IsNullOrEmpty(TextBox1.Text))
            answercountlow = Convert.ToInt32(TextBox1.Text);
        int answercountupp = 0;
        if (!string.IsNullOrEmpty(TextBox2.Text))
            answercountupp = Convert.ToInt32(TextBox2.Text);
        string userName = "";
        if (TextBoxMember.Text != "") userName = TextBoxMember.Text;
        string email = "";
       // if (TextBoxEMail.Text != "") email = TextBoxEMail.Text;
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

        IList historyTop = historyPicture.GetTopUnsys(userName, email, scoreLowerBound, scoreUpper, answercountlow, answercountupp, startTime, endTime, pageSize, pageNumber);
        StringBuilder sb = new StringBuilder();
        sb.Append("<table class=\"type02\" width=\"60%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" >");
        sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
        sb.Append("<th scope=\"col\" width=\"5%\">停權</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">帳號</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">活動得分</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">答題數</th></tr>");
        if (historyTop.Count > 0)
        {
            for (int i = 0; i < historyTop.Count; i++)
            {
                sb.Append("<tr><td>" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                HistoryPicture.HistoryTop topObj = (HistoryPicture.HistoryTop)historyTop[i];
                if (topObj.disable == true)
                {
                    sb.Append("<td align=\"center\"><input type=\"checkbox\" checked=\"checked\" name=\"unDisableUser\" value=\"" + topObj.LoginId + "\" /></td>");
                }
                else
                {
                    sb.Append("<td align=\"center\"><input type=\"checkbox\" name=\"DisableUser\" value=\"" + topObj.LoginId + "\" /></td>");
                }
                sb.Append("<td align=\"center\">" + "<a href=\"" + kmwebsysSite + "/kmactivity/history/historylog.aspx?type=all&accountid=" + topObj.AccountId.ToString() + "\" >" + topObj.LoginId + "</a></td>");
                if (!string.IsNullOrEmpty(topObj.NickName))
                {
                    sb.Append("<td align=\"center\">" + topObj.NickName + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + (topObj.RealName.Substring(0, 1) + "＊" + topObj.RealName.Substring(topObj.RealName.Trim().Length - 1, 1)) + "</td>");
                }
                else
                {
                    sb.Append("<td>&nbsp;</td>");
                }
                sb.Append("<td align=\"center\">" + topObj.Score + "</td>");
                sb.Append("<td align=\"center\">" + topObj.Count + "</td>");
            }
        }
        int totaCount = historyPicture.GetTopCounUnsys(userName, email, scoreLowerBound, scoreUpper, answercountlow, answercountupp, startTime, endTime);
        TotalRecordText.Text = totaCount.ToString();
        int pageCount = Convert.ToInt32((totaCount / pageSize + 0.999));
        if ((totaCount % pageSize) == 0)
            pageCount = Convert.ToInt32((totaCount / pageSize));
        TotalPageText.Text = pageCount.ToString();
        ListItem item = default(ListItem);
        PageNumberDDL.Items.Clear();
        for (int j = 0; j <= pageCount-1; j++)
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
            PreviousLink.NavigateUrl = "activityrankdetail.aspx?PageNumber=" + (pageNumber - 1).ToString() + "&PageSize=" + pageSize.ToString();
        }
        else
        {
            PreviousText.Enabled = false;
            PreviousLink.NavigateUrl = "";
        }
        if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
        {
            NextLink.NavigateUrl = "activityrankdetail.aspx?PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString();
        }
        else
        {
            NextText.Enabled = false;
            NextLink.NavigateUrl = "";
        }
        sb.Append("</table>");
        TableText.Text = sb.ToString();
        string urlTemp = kmwebsysSite + "/kmactivity/history/activityrankdetalexport.aspx?querymember=" + HttpUtility.UrlEncode(userName);
        urlTemp += "&scoreUpper=" + scoreUpper + "&scoreLowerBound=" + scoreLowerBound + "&answercountlow=" + answercountlow.ToString() + "&answercountupp=" + answercountupp.ToString();
        urlTemp += "&starttime=" + startTime + "&endtime=" + endTime ;
        linkExport.NavigateUrl = urlTemp;
    }

    private void UpdateDisableUser(int state, string login_ID)
    {
        string sql = @"
            update ACCOUNT set Disable = @Disable where login_id = @login_id
            update mGIPcoanew..member set status = @status  where account = @login_id
        ";
        string status = "Y";
        if (state == 1)
            status = "N";
        string[] ss = login_ID.Split(',');
        foreach (string s in ss)
        {
            SqlHelper.ExecuteNonQuery("HistoryPictureConnString", sql,
                DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", s),
                 DbProviderFactories.CreateParameter("HistoryPictureConnString", "@Disable", "@Disable", state),
                  DbProviderFactories.CreateParameter("HistoryPictureConnString", "@status", "@status", status));
        }
    }
}