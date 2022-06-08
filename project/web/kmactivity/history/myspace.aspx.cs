using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections;
using GSS.Vitals.COA.Data;
using System.Data;

public partial class kmactivity_history_myspace : System.Web.UI.Page
{
    private HistoryPicture historyPicture;
    protected void Page_Load(object sender, EventArgs e)
    {

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
            Response.StatusCode = 403;
            Response.End();
        }
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        string loginId = Session["memID"].ToString().Trim();
        historyPicture = new HistoryPicture(loginId);
        historyPicture.loginUpdate(loginId);
        HistoryPicture.HistoryTop historyTop = historyPicture.GetUserScoreData(loginId);
        int userDailyAnswer = historyPicture.GetUserDailyAnswer(loginId);
        UserDailyAnswer.Text = userDailyAnswer.ToString();
        if (historyTop != null)
        {
            DisplayData(historyTop, loginId);
        }
        else
        {
            DisplayDefaultData(loginId);
        }
    }

    private void DisplayDefaultData(string loginId)
    {
        string sql = "select realname,nickname from member where account = @account";
        var dt = SqlHelper.GetDataTable("GSSConnString",sql,
            DbProviderFactories.CreateParameter("GSSConnString", "@account", "@account", loginId));
        DataRow dr = dt.Rows[0];
        userName.Text = dr["nickname"].ToString();
        loginid.Text = loginId;
        UsersTotalScore.Text = "0";
        AllQuestion.Text = "0";
        CurrentQuestion.Text = "0";
        //DisplayQuestionList(loginId);
    }

    private void DisplayData(HistoryPicture.HistoryTop historyTop, string loginId)
    {
        userName.Text = historyTop.NickName;
        loginid.Text = loginId;
        UsersTotalScore.Text = historyTop.Score.ToString();
        AllQuestion.Text = historyPicture.GerUserQuestioninfoCount(loginId).ToString();
        CurrentQuestion.Text = historyTop.Count.ToString();
        DisplayQuestionList(loginId);
    }

    private void DisplayQuestionList(string loginId)
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
            pageSize =15;
            pageNumber = Convert.ToInt32(PageNumberDDL.SelectedValue);
        }
        IList objlist = historyPicture.GerUserDailyScoreinfo(loginId, pageSize, pageNumber);
        string ss = "<table width=\"100%\" id=\"rank\"><tr><th></th><th>日期</th><th>答題數</th><th>得分</th></tr>";
        foreach (HistoryPicture.HistoryQuestionInfo obj in objlist)
        {
            ss += "<tr>";
            ss += "<td>" + obj.RowNumber + "</td>";
            ss += "<td>" + obj.CreateDate + "</td>";
            ss += "<td>" + obj.Icuitem + "</td>";
            ss += "<td>" + obj.Score + "</td>";
            ss += "</tr>";
        }
        ss += "</table>";
        treasureTable.Text = ss;
        PageNumberText.Text = pageNumber.ToString();
        int totaCount = historyPicture.GerUserDailyScoreinfoCount(loginId);
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
            PreviousLink.NavigateUrl = "myspace.aspx?PageNumber=" + (pageNumber - 1).ToString() + "&PageSize=" + pageSize.ToString() ;
        }
        else
        {
            PreviousText.Enabled = false;
            PreviousLink.NavigateUrl = "";
        }
        if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
        {
            NextLink.NavigateUrl = "myspace.aspx?PageNumber=" + (pageNumber + 1).ToString() + "&PageSize=" + pageSize.ToString();
        }
        else
        {
            NextText.Enabled = false;
            NextLink.NavigateUrl = "";
        }
        TotalRecordText.Text = totaCount.ToString();
    }

}