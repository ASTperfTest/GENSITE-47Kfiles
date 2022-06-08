using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections;
using GSS.Vitals.COA.Data;

public partial class kmactivity_history_historylog : System.Web.UI.Page
{
    private HistoryPicture historyPicture;
    protected string account_id = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        string accountId = WebUtility.GetStringParameter("accountid", string.Empty).ToLower();
        historyPicture = new HistoryPicture("");
        DisplayQuestionList(accountId);
        account_id = accountId.ToString();
    }

    private void DisplayQuestionList(string accountId)
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
        string sql = @" select * from account where account_id = @account_id
        ";
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@account_id", "@account_id", accountId));
        string loginId = dt.Rows[0]["login_id"].ToString();
        IList objlist = historyPicture.GerUserQuestioninfo(loginId, pageSize, pageNumber);
        loginid.Text = loginId;
        userName.Text = dt.Rows[0]["realname"].ToString();
        userMail.Text = dt.Rows[0]["email"].ToString();
        string ss = "<table width=\"50%\"><tr><th></th><th>日期</th><th>答案</th><th>答題狀況</th><th>得分</th><th>縮圖</th></tr>";
        foreach (HistoryPicture.HistoryQuestionInfo obj in objlist)
        {
            ss += "<tr>";
            ss += "<td>" + obj.RowNumber + "</td>";
            ss += "<td>" + obj.CreateDate + "</td>";
            ss += "<td>" + obj.Stitle + "</td>";
            if (obj.State == 3)
            {
                ss += "<td>答題成功</td>";
            }
            else
            {
                ss += "<td>答題失敗</td>";
            }
            string imageUrl = "<img src=\"/public/History/" + historyPicture.GetCurrentPic(obj.Icuitem) + "\" width=\"30px\" height=\"30px\" /> ";

            ss += "<td>" + obj.Score + "</td>";
            ss += "<td>" + imageUrl + "</td>";
            ss += "</tr>";
        }
        ss += "</table>";
        TableText.Text = ss;
        PageNumberText.Text = pageNumber.ToString();
        int totaCount = historyPicture.GerUserQuestioninfoCount(loginId);
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
            PreviousLink.NavigateUrl = "historylog.aspx?PageNumber=" + (pageNumber - 1).ToString() + "&PageSize=" + pageSize.ToString() + "&accountid=" + accountId.ToString();
        }
        else
        {
            PreviousText.Enabled = false;
            PreviousLink.NavigateUrl = "";
        }
        if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
        {
            NextLink.NavigateUrl = "historylog.aspx?PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString() + "&accountid=" +accountId.ToString();
        }
        else
        {
            NextText.Enabled = false;
            NextLink.NavigateUrl = "";
        }
        TotalRecordText.Text = totaCount.ToString();
    }
}