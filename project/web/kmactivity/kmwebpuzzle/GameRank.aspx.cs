using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Transactions;
using System.Data.Linq;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using GSS.Vitals.COA.Data;

// 基本功能應該都已經OK了
public partial class GameRank : System.Web.UI.Page
{
    private PagedDataSource Pager;
    private DataTable dt;
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            GetData(0, 15);
        }
    }

    protected void GetData(int pageNumber, int pageSize) 
    {
        string sql = @"
        

        select  ROW_NUMBER() OVER(order by TotalGrade desc,B.Counting,C.useenergy)as Row,GH.LOGIN_ID,
        isnull(B.Counting,0) as Counting,isnull(B.TotalGrade,0) as TotalGrade,isnull(C.useenergy,0) as useenergy,  mm.realname,isnull(mm.nickname,'') as nickname,mm.email  from (
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

                 from GAMEHistory where ( picState='N' or picState='Y' ) {0}
                group by LOGIN_ID,pic_id
        ) G 
        group by LOGIN_ID
                ) C on C.login_id = GH.login_id 
left  join mGIPcoanew..member mm on mm.account = b.login_id
                where ( mm.status <> 'N' or mm.status is null ) and TotalGrade >0
          

        ";

        string gametimetemp = " and convert(nvarchar,gametime,111) < convert(nvarchar,dateadd(\"d\",1,getdate()),111) ";
        sql = string.Format(sql, gametimetemp);
        dt = SqlHelper.GetDataTable("PuzzleConnString", sql.ToString());
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
            if (realname.ToString().Length >= 3)
            {
                str = realname.ToString().Substring(0, 1) + "＊" + realname.ToString().Substring(2, realname.ToString().Length - 2);
                return str;
            }
            else
            {
                return realname.ToString().Substring(1, 1) + "＊";
            }
        }
        
    }

}

