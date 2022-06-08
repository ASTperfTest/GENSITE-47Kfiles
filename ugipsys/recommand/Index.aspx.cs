using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Data;
using System.Transactions;

/// <summary>
/// Added By Leo    2011-06-24      報馬仔文章管理
/// </summary>
public partial class Index : System.Web.UI.Page
{
    private PagedDataSource Pager;
    private DataTable dt;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            int PageNumber = 0;
            int PageSize = 10;

            myDBinit(PageNumber, PageSize);
        }
    }


    protected void myDBinit(int intPageNumber, int intPageSize)
    {
        string sqlQueryScript = @"
        SELECT 
	         RecommandContent.* 
	        ,member.account
	        ,member.realname
	        ,member.nickname
        FROM RecommandContent
        inner join member on member.account = RecommandContent.iEditor";

        string strStatus = string.Empty;
        string strKeywords = string.Empty;

        // 審核狀態
        switch (ddlStatus.SelectedValue)
        {
            case "1":
                strStatus = " RecommandContent.Status IS NULL ";
                break;
            case "2":
                strStatus = " RecommandContent.Status = 'y' ";
                break;
            case "3":
                strStatus = " RecommandContent.Status = 'n' ";
                break;
            default:
                break;
        }

        // 關鍵字
        if(!string.IsNullOrEmpty(txtKeyword.Text.Trim()))
        {
            strKeywords = " Title LIKE '%" + txtKeyword.Text.Trim() + "%' ";
        }

        // 組SQL Script
        if (!string.IsNullOrEmpty(strStatus) && !string.IsNullOrEmpty(strKeywords))
        {
            sqlQueryScript += " WHERE " + strStatus + " AND " + strKeywords;
        }
        else if (!string.IsNullOrEmpty(strStatus))
        {
            sqlQueryScript += " WHERE " + strStatus;
        }
        else if(!string.IsNullOrEmpty(strKeywords))
        {
            sqlQueryScript += " WHERE " + strKeywords;
        }
        sqlQueryScript += " ORDER BY created_Date DESC";


        //Response.Write(sqlQueryScript);
        
        dt = SqlHelper.GetDataTable("ConnString", sqlQueryScript);

        Pager = dt.Paging(intPageNumber, intPageSize);
        rptList.DataSource = Pager;
        rptList.DataBind();

        SetControl();
    }

    // 設置控制項
    protected void SetControl()
    {
        // 設定PageNumber_DropDownList
        PageNumberDDL.Items.Clear();
        for (int i = 0; i < Pager.PageCount; i++)
        {
            PageNumberDDL.Items.Add(new ListItem((i + 1).ToString(), i.ToString()));
        }
        PageNumberDDL.SelectedIndex = Pager.CurrentPageIndex;

        // 上、下頁的Visible
        if (PageNumberDDL.Items.Count > 1)
        {
            if (Pager.IsFirstPage)
            {
                PreviousLink.Visible = false;
                NextLink.Visible = true;
            }
            else if (Pager.IsLastPage)
            {
                PreviousLink.Visible = true;
                NextLink.Visible = false;
            }
            else
            {
                PreviousLink.Visible = true;
                NextLink.Visible = true;
            }
        }
        else
        {
            PreviousLink.Visible = false;
            NextLink.Visible = false;
        }

        PageNumberText.Text = (Pager.CurrentPageIndex + 1).ToString();
        TotalPageText.Text = Pager.PageCount.ToString();
        TotalRecordText.Text = dt.Rows.Count.ToString();
    }

    protected void PageNumberDDL_SelectedIndexChanged(object sender, EventArgs e)
    {
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    protected void PageSizeDDL_SelectedIndexChanged(object sender, EventArgs e)
    {
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    protected void PreviousLink_Click(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex--;
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    protected void NextLink_Click(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex++;
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
      
    // Repeater ItemDataBound事件
    protected void rptList_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            string cID = ((Label)e.Item.FindControl("labcID")).Text;
            // 審核
            switch (((Label)e.Item.FindControl("labStatus")).Text.ToUpper())
            {
                case "Y":
                    ((Label)e.Item.FindControl("labStatus")).Text = "<font color='blue'>通過</font>";
                    break;
                case "N":
                    ((Label)e.Item.FindControl("labStatus")).Text = "<font color='red'>未通過</font>";
                    break;
                default:
                    ((Label)e.Item.FindControl("labStatus")).Text = "尚未審核";
                    break;
            }

            // 推薦日期
            DateTime dt = Convert.ToDateTime(((Label)e.Item.FindControl("labDate")).Text);
            ((Label)e.Item.FindControl("labDate")).Text = string.Format("{0:yyyy/MM/dd}", dt);

            // TAGs
            string sqlQueryTAGsNameScript = @"SELECT B.tagName 
                                              FROM RecommandContent2TAGs AS A INNER JOIN TAGs AS B
                                                    ON A.tagID = B.tagID 
                                              WHERE A.cid = @cID";
            using (var reader = SqlHelper.ReturnReader("ConnString", sqlQueryTAGsNameScript,
                DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", cID)))
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        ((Label)e.Item.FindControl("labTAGs")).Text += reader["tagName"] + ",";
                    }
                }
            }

            // 編修網址
            ((HyperLink)e.Item.FindControl("HyperLink1")).NavigateUrl = "Article_Edit.aspx?cID=" + cID;
            //((HyperLink)e.Item.FindControl("HyperLink1")).Target = "_blank";
        }

    }

    // 通過
    protected void btnYes_Click(object sender, EventArgs e)
    {
        // Using TransactionScop
        using (TransactionScope Scope = new TransactionScope())
        {
            try
            {
                // 查詢報馬仔的配分
                string sqlSharePointScript = @"SELECT Rank0_1 FROM kpi_set_score WHERE Rank0 = 'st_4' AND Rank0_2 = 'st_420'";
                string strSharePoint = SqlHelper.ReturnScalar("ConnString", sqlSharePointScript).ToString();

                for (int i = 0; i < rptList.Items.Count; i++)
                {
                    if (((CheckBox)rptList.Items[i].FindControl("checkbox1")).Checked)
                    {
                        int cID = Convert.ToInt32(((Label)rptList.Items[i].FindControl("labcID")).Text);
                        
                        sqlSharePointScript = @"
                            declare @recommand int set @recommand = @cID --@cID由程式代入
                            declare @Rank01 int
                            set @Rank01 = isnull((SELECT top 1 Rank0_1 FROM kpi_set_score WHERE Rank0 = 'st_4' AND Rank0_2 = 'st_420'), 0)

                            declare @member varchar(30)
                            declare @contentDate datetime
                            declare @contentDateTime datetime
                            declare @kpiSNO int
                            declare @shareSNO int

                            select 
	                             @contentDate=convert(varchar,Created_Date,111)
                                ,@contentDateTime=Created_Date
	                            ,@kpiSNO=kpisno
	                            ,@member=iEditor
                            from RecommandContent where cID = @recommand and ([Status] != 'Y' or [Status] is null)

                            --同意推薦，增加KPI分數
                            if not @contentDate is null
                            begin
                            	
	                            --查詢KPI記錄是否存在
	                            set @shareSNO = (select max(sno) from MemberGradeShare where memberId=@member and (shareDate>=@contentDate and shareDate<@contentDate+1))
	                            if @shareSNO is null
	                            begin
		                            insert into MemberGradeShare (memberId, shareDate) select  @member ,@contentDateTime
		                            set @shareSNO = SCOPE_IDENTITY()
	                            end
                            	
	                            update MemberGradeShare 
	                            set shareRecommend=shareRecommend+@Rank01 
	                            where memberId=@member and sno=@shareSNO	
                            	
	                            update RecommandContent 
	                            set   KPIPoint = @Rank01
		                            , KPISNO = @shareSNO 
		                            , [Status] = 'Y'
	                            where cID = @recommand
                            end";

                        SqlHelper.ExecuteNonQuery("ConnString", sqlSharePointScript,
                            DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", cID));
                        AddPuzzleScore(cID);
                    }
                }
                myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
                Scope.Complete();

            }
            catch (Exception)
            {
                throw;
            }
        }
    }

    // 未通過
    protected void btnNo_Click(object sender, EventArgs e)
    {
        // Using TransactionScop
        using (TransactionScope Scope = new TransactionScope())
        {
            try
            {
                // 查詢報馬仔的配分
                string sqlSharePointScript = @"SELECT Rank0_1 FROM kpi_set_score WHERE Rank0 = 'st_4' AND Rank0_2 = 'st_420'";
                string strSharePoint = SqlHelper.ReturnScalar("ConnString", sqlSharePointScript).ToString();
                for (int i = 0; i < rptList.Items.Count; i++)
                {
                    if (((CheckBox)rptList.Items[i].FindControl("checkbox1")).Checked)
                    {
                        int cID = Convert.ToInt32(((Label)rptList.Items[i].FindControl("labcID")).Text);

                        sqlSharePointScript = @"
                            declare @recommand int set @recommand = @cID --@cID由程式代入

                            declare @member varchar(30)
                            declare @contentDate datetime
                            declare @kpiSNO int
                            declare @KPIPoint int
                            declare @shareSNO int

                            select 
	                             @contentDate=convert(varchar,Created_Date,111)
	                            ,@kpiSNO=kpisno
	                            ,@KPIPoint=KPIPoint
	                            ,@member=iEditor
                            from RecommandContent where cID = @recommand and ([Status] = 'Y' or [Status] is null)

                            --不同意推薦，增加KPI分數(如果之前有加KPI)
                            if not @contentDate is null
                            begin
                            		
	                            update MemberGradeShare 
	                            set shareRecommend=shareRecommend - isnull(@KPIPoint, 0)
	                            where memberId=@member and sno=@kpiSNO and @kpiSNO>0
                            	
	                            update RecommandContent 
	                            set   KPIPoint = 0
		                            , KPISNO = 0
		                            , [Status] = 'N'
	                            where cID = @recommand
                            end";

                        SqlHelper.ExecuteNonQuery("ConnString", sqlSharePointScript,
                            DbProviderFactories.CreateParameter("ConnString", "@cID", "@cID", cID));
                        
                    }
                }
                myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
                Scope.Complete();
            }
            catch (Exception)
            {
                throw;
            }
        }
    }

    protected void ddlStatus_SelectedIndexChanged(object sender, EventArgs e)
    {
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    protected void btnSearch_Click(object sender, EventArgs e)
    {
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }

    private void AddPuzzleScore(int cid)
    {
        string sql = @"
        
        DECLARE @login_id nvarchar(50);
        set  @login_id = (select iEditor from RecommandContent where cid = @cid );
        if not exists(
        select * from AddPuzzlePouint ap 
        left join RecommandContent r on r.iEditor = ap.account
        where   CONVERT(varchar(12) , r.created_date , 111 ) = CONVERT(varchar(12) , ap.createdate , 111 )
        and r.cid = @cid
        )
        begin
            if not exists( select * from Puzzle..ACCOUNT where login_id =  @login_id)
	        begin
	        insert into Puzzle..ACCOUNT (LOGIN_ID,REALNAME,NICKNAME,EMAIL,Energy,CREATE_DATE,MODIFY_DATE)
	        select account,REALNAME,isnull(NICKNAME,''),EMAIL,50,getdate(),getdate() from member where account =  @login_id
	        end
	        else
	        begin
	        update Puzzle..ACCOUNT set REALNAME = m.REALNAME,NICKNAME=isnull(m.NICKNAME,''),email = m.email,Energy=Energy+50,MODIFY_DATE=getdate()
	        from (select * from member  where account =  @login_id ) m
	        where Puzzle..ACCOUNT.login_id =  @login_id
	        end
            insert into AddPuzzlePouint (account,createdate)
            values(@login_id,getdate())

        end
        
        ";
        SqlHelper.ExecuteNonQuery("ConnString", sql,
            DbProviderFactories.CreateParameter("ConnString", "@cid", "@cid", cid));
    }

}