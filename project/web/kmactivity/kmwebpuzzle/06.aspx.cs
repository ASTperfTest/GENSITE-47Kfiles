using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Data;

public partial class kmactivity_kmwebpuzzle_05 : System.Web.UI.Page
{
    string meid = "";
    private PagedDataSource Pager;
    protected void Page_Load(object sender, EventArgs e)
    {
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        if (Session["memID"] == null || Session["memID"].ToString() == "")
        {
            Response.Write("<script>alert('連線逾時或尚未登入，請登入會員!');window.location.href=\"01.aspx?a=puzzle\";</script>");
            Response.End();
        }

        meid = Session["memID"].ToString().Trim();
        if (!IsPostBack)
        {
            DisplayUserData( 0);
        }

    }

    private void DisplayUserData(int pageNumber)
    {
        string sql = @"
        select  B.LOGIN_ID,B.Energy,B.REALNAME, B.NICKNAME, B.Counting, (B.easy*1+(B.Counting-B.easy)*2) as TotalGrade 
        from( select A.LOGIN_ID,A.Energy,A.REALNAME, A.NICKNAME,A.Counting, count(difficult) as easy 
              from ( select ACCOUNT.LOGIN_ID,ACCOUNT.Energy, ACCOUNT.REALNAME, ACCOUNT.NICKNAME, count(picState)as Counting 
                     from ACCOUNT 
                     left join GAMEHistory on ACCOUNT.LOGIN_ID = GAMEHistory.LOGIN_ID 
                     where picState='Y' 
                     group by ACCOUNT.LOGIN_ID, ACCOUNT.REALNAME, ACCOUNT.NICKNAME, picState ,ACCOUNT.Energy
                    ) as A 
              left join GAMEHistory on A.LOGIN_ID = GAMEHistory.LOGIN_ID 
              where difficult='E' and picState='Y' 
              group by A.LOGIN_ID,A.REALNAME, A.NICKNAME, A.Counting ,A.Energy
        ) as B 
        where login_id = @login_id
        ";
        var dt = SqlHelper.GetDataTable("PuzzleConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", meid));
        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            loginid.Text = dr["LOGIN_ID"].ToString();
            userEnergy.Text = dr["Energy"].ToString();
            UsersTotalScore.Text = dr["TotalGrade"].ToString();
            AllQuestion.Text = dr["Counting"].ToString();
            sql = @"
                select aa.login_id,isnull(A.dailycounting,0) as dailycounting,isnull(B.dailyfcounting,0) as dailyfcounting from 
                account aa left join 
                (
                select login_id,count(*) as dailycounting from dbo.GAMEHistory where login_id = @login_id
                                and convert(varchar,gametime,111) = convert(varchar,getdate(),111)
                                and picState = 'Y'
				                group by login_id
                ) A on A.login_id = aa.login_id
                left join 
                (
                select login_id,count(*) as dailyfcounting from dbo.GAMEHistory where login_id = @login_id
                                and convert(varchar,gametime,111) = convert(varchar,getdate(),111)
                                and picState = 'F'
				                group by login_id

                ) B on aa.login_id = b.login_id
                where aa.login_id = @login_id
            ";
            dt.Dispose();
            dt = SqlHelper.GetDataTable("PuzzleConnString", sql,
                DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", meid));
            if (dt.Rows.Count > 0)
            {
                dr = dt.Rows[0];

                UserDailyAnswer.Text = dr["dailycounting"].ToString();
                UserDailyFAnswer.Text = dr["dailyfcounting"].ToString();
            }
        }
        else
        {
            dt.Dispose();
            sql = "select * from ACCOUNT where login_id = @login_id ";
            dt = SqlHelper.GetDataTable("PuzzleConnString", sql,
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", meid));
            if (dt.Rows.Count > 0)
            {
                DataRow dr = dt.Rows[0];
                loginid.Text = dr["LOGIN_ID"].ToString();
                userEnergy.Text = dr["Energy"].ToString();
            }
            else
            {
                loginid.Text = meid;
            }
            
        }
        dt.Dispose();

        sql = @"
     
        select  ROW_NUMBER() OVER(order by GH.gametime desc)as Row,GH.LOGIN_ID,GH.gametime,
        isnull(B.Counting,0) as Counting,isnull(B.TotalGrade,0) as TotalGrade,isnull(C.FCounting,0) as FCounting  from (
        select login_id,convert(nvarchar,gametime,111) as gametime from  GAMEHistory 
        where ( picState='F' or picState='Y' ) and login_id = @login_id
        group by convert(nvarchar,gametime,111),login_id 
        ) GH
        left join 
        (
	        select  LOGIN_ID,gametime,sum(Counting) as Counting,sum (score) as  TotalGrade
                from (
                select LOGIN_ID,convert(nvarchar,gametime,111) as gametime,count(picState) as Counting,difficult,case difficult when 'E' then 1 * count(picState) else 2 * count(picState) end as score

                 from GAMEHistory where picState='Y' 
		        and  login_id = @login_id
                group by convert(nvarchar,gametime,111),LOGIN_ID,difficult
                ) A 
                group by gametime,LOGIN_ID 
        ) B on B.LOGIN_ID = GH.LOGIN_ID and B.gametime = GH.gametime
        left  join 
                (
	                 select  LOGIN_ID,gametime,sum(Counting) as FCounting
                        from (
                        select LOGIN_ID,convert(nvarchar,gametime,111) as gametime,count(picState) as Counting

                         from GAMEHistory where picState='F' 
		                and  login_id = @login_id
                        group by convert(nvarchar,gametime,111),LOGIN_ID,difficult
                        ) A 
                        group by gametime,LOGIN_ID 
                ) C on C.login_id = GH.login_id and c.gametime = GH.gametime
        ";
        dt = SqlHelper.GetDataTable("PuzzleConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", meid));
        Pager = dt.Paging(pageNumber, 15);
        rpList.DataSource = Pager;
        rpList.DataBind();
        SetControl(dt);
    }
    protected void SetControl(DataTable dt)
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
    protected void preLinkAct(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex--;
        DisplayUserData(Convert.ToInt32(PageNumberDDL.SelectedValue));
    }
    protected void nextLinkAct(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex++;
        DisplayUserData(Convert.ToInt32(PageNumberDDL.SelectedValue));
    }
    protected void ChangePageNumber(object sender, EventArgs e)
    {
        DisplayUserData(Convert.ToInt32(PageNumberDDL.SelectedValue));
    }

}