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
using System.Transactions;
using System.Data.Linq;
using System.Data.SqlClient;
using System.Data;
using GSS.Vitals.COA.Data;

public partial class BackstageGameRank : System.Web.UI.Page
{
    private string kmwebsysSite = ConfigurationManager.AppSettings["myURL"].ToString();
    private PagedDataSource Pager;
    private DataTable dt;
    protected void Page_Load(object sender, EventArgs e)
    {
        string a = WebUtility.GetStringParameter("a", "");
        if (a != "edrftg")
        {
            Response.StatusCode = 403;
            Response.End();
        }
        if (!IsPostBack)
        {
            GetData(0, 15);
        }
    }

    protected void CheckCheat(object sender, EventArgs e) 
    {
        string urlTemp = kmwebsysSite + "/kmactivity/kmwebpuzzle/CheckCheatUser.aspx";
        Response.Redirect(urlTemp);
    }
    protected void GetData(int pageNumber, int pageSize) 
    {
        string sql = @"
            select  ROW_NUMBER() OVER(order by TotalGrade desc,B.Counting,C.useenergy)as Row,GH.LOGIN_ID,
        isnull(B.Counting,0) as Counting,isnull(B.TotalGrade,0) as TotalGrade,
		isnull(C.useenergy,0) as useenergy,  mm.realname,isnull(mm.nickname,'') as nickname,mm.email,ac.Energy,ac.GetEnergy,
        case(isnull(mm.status,'Y') ) when 'Y' then 0 else 1 end as status from (
        select login_id from  GAMEHistory 
        where ( picState='F' or picState='Y' ) {0}
        group by login_id 
        ) GH
        left join 
        (
	        select  LOGIN_ID,sum(Counting) as Counting,sum (score) as  TotalGrade
                from (
                select LOGIN_ID,count(picState) as Counting,difficult,case difficult when 'E' then 1 * count(picState) else 2 * count(picState) end as score

                 from GAMEHistory where picState='Y' {0}
		        
                group by LOGIN_ID,difficult
                ) A 
                group by LOGIN_ID 
        ) B on B.LOGIN_ID = GH.LOGIN_ID 
        left  join 
                (
	                 select LOGIN_ID,sum(useenergy)-count(pic_id) as useenergy from (
        select LOGIN_ID,pic_id,count(picState) as useenergy

                 from GAMEHistory where  ( picState='N' or picState='Y' ) {0}
                group by LOGIN_ID,pic_id
        ) G 
        group by LOGIN_ID
                ) C on C.login_id = GH.login_id 
left  join ACCOUNT ac on ac.login_id = GH.login_id
left  join mGIPcoanew..member mm on mm.account = GH.login_id
where 1=1

         ";
        if (TextBoxMember.Text != "")
        {
            sql += " and  (mm.account like @login_id or mm.REALNAME like @login_id or mm.NICKNAME like @login_id) ";
        }
        int scoreUpper = -1;
        if (!string.IsNullOrEmpty(TextScoreUpperBound.Text))
            scoreUpper = Convert.ToInt32(TextScoreUpperBound.Text);
        int scoreLowerBound = -1;
        if (!string.IsNullOrEmpty(TextScoreLowerBound.Text))
            scoreLowerBound = Convert.ToInt32(TextScoreLowerBound.Text);
        if (scoreUpper != -1 && scoreLowerBound != -1)
        {
            sql += " and TotalGrade between @scoreLowerBound and @scoreUpper ";
        }
        int answercountlow = -1;
        if (!string.IsNullOrEmpty(TextBox1.Text))
            answercountlow = Convert.ToInt32(TextBox1.Text);
        int answercountupp = -1;
        if (!string.IsNullOrEmpty(TextBox2.Text))
            answercountupp = Convert.ToInt32(TextBox2.Text);
        if (answercountlow != -1 && answercountupp != -1)
        {
            sql += " and Counting between @answercountlow and @answercountupp  ";
        }
        string startTime = "";
        string endTime = "";
        DateTime timeTemp = new DateTime();
        string timestring = "";
        if (TxtStartDate.Text != "" && TxtEndDate.Text != "")
        {
            if (DateTime.TryParseExact(TxtStartDate.Text, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp) && DateTime.TryParseExact(TxtEndDate.Text, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp))
            {
                startTime = TxtStartDate.Text;
                endTime = TxtEndDate.Text;
                timestring = " and gametime between @startTime and dateadd(\"d\",1,@endTime)  ";
            }
        }
        string urlTemp = kmwebsysSite + "/kmactivity/kmwebpuzzle/GameRankExport.aspx?querymember=" + HttpUtility.UrlEncode(TextBoxMember.Text.ToString().Trim());
        urlTemp += "&scoreUpper=" + scoreUpper + "&scoreLowerBound=" + scoreLowerBound + "&answercountlow=" + answercountlow.ToString() + "&answercountupp=" + answercountupp.ToString();
        urlTemp += "&starttime=" + startTime + "&endtime=" + endTime;
        linkExport.NavigateUrl = urlTemp;
        sql = string.Format(sql, timestring);
        dt = SqlHelper.GetDataTable("PuzzleConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id","%"+ TextBoxMember.Text.ToString().Trim() +"%"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scoreLowerBound", "@scoreLowerBound", scoreLowerBound),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scoreUpper", "@scoreUpper", scoreUpper),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountlow", "@answercountlow", answercountlow),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountupp", "@answercountupp", answercountupp),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@startTime", "@startTime", startTime),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@endTime", "@endTime", endTime));
        Pager = dt.Paging(pageNumber, pageSize);
        rpList.DataSource = Pager;
        rpList.DataBind();
        SetControl();
    }

    protected void preLinkAct(object sender, EventArgs e) 
    {
        PageNumberDDL.SelectedIndex--;
        GetData(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
    protected void nextLinkAct(object sender, EventArgs e) 
    {
        PageNumberDDL.SelectedIndex++;
        GetData(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
    protected void ChangePageNumber(object sender, EventArgs e)
    {
        GetData(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
    protected void ChangePageSize(object sender, EventArgs e) 
    {
        GetData(0, Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
    protected void SetControl()
    {
        PageNumberDDL.Items.Clear();
        for (int i = 0; i < Pager.PageCount; i++)
        {
            PageNumberDDL.Items.Add(new ListItem((i + 1).ToString(), i.ToString()));
        }
        PageNumberDDL.SelectedIndex = Pager.CurrentPageIndex;

        if (PageNumberDDL.Items.Count > 1)
        {
            if (Pager.IsFirstPage)
            {
                preLink.Visible = false;
                nextLink.Visible = true;
            }
            else if (Pager.IsLastPage)
            {
                preLink.Visible = true;
                nextLink.Visible = false;
            }
            else
            {
                preLink.Visible = true;
                nextLink.Visible = true;
            }
        }
        else
        {
            preLink.Visible = false;
            nextLink.Visible = false;
        }
        PageNumberText.Text = (Pager.CurrentPageIndex + 1).ToString();
        TotalPageText.Text = Pager.PageCount.ToString();
        TotalRecordText.Text = dt.Rows.Count.ToString();
    }

    protected string DealName(object realname, object nickname)
    {
        string str = "";
        if (nickname != null && !string.IsNullOrEmpty(nickname.ToString()))
        {
            return nickname.ToString();
        }
        else {
            str = realname.ToString();
            return str;
        }

    }

    private void ChangeUserStates()
    {
        string disalbeUser = "";
        string enableUser = "";
        for (int i = 0; i < rpList.Items.Count; i++)
        {
            if (((CheckBox)rpList.Items[i].FindControl("CheckBox1")).Checked)
            {
                if (disalbeUser == "")
                {
                    disalbeUser += ((Label)rpList.Items[i].FindControl("labmemberid")).Text;
                }
                else
                {
                    disalbeUser += "," + ((Label)rpList.Items[i].FindControl("labmemberid")).Text;
                }
            }
            else
            {
                if (enableUser == "")
                {
                    enableUser += ((Label)rpList.Items[i].FindControl("labmemberid")).Text;
                }else
                {
                    enableUser += "," + ((Label)rpList.Items[i].FindControl("labmemberid")).Text;
                }
            }
        }
        
        if(disalbeUser != "")
            UpdateDisableUser(1,disalbeUser);
        if (enableUser != "")
            UpdateDisableUser(0,enableUser);
        GetData(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    private void UpdateDisableUser(int state, string login_ID)
    {
        string sql = @"
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
                  DbProviderFactories.CreateParameter("HistoryPictureConnString", "@status", "@status", status));
        }
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        GetData(0, 15);
    }

    protected void Button2_Click(object sender, EventArgs e)
    {
        ChangeUserStates();
        GetData(int.Parse(PageNumberDDL.Text)-1, 15);
    }
}