using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using GSS.Vitals.COA.Data;
using System.Configuration;

public partial class kmactivity_kmwebpuzzle_UserGameLog : System.Web.UI.Page
{
    private PagedDataSource Pager;
    private string login_id = "";
    private DataTable dt;
    private string kmwebsysSite = ConfigurationManager.AppSettings["myURL"].ToString();
    protected void Page_Load(object sender, EventArgs e)
    {
        login_id = WebUtility.GetStringParameter("loginid", string.Empty).ToLower();
        if (!IsPostBack)
        {
            DisplayData(0,15);
        }
    }

    private void DisplayData(int pageNumber, int pageSize)
    {
        string sql = @"
            select ROW_NUMBER() OVER(order by gametime DESC) as Row , login_id,
             picstate,GAMEHistory.pic_id,convert(nvarchar,gametime,111) as gametime,difficult,pic.pic_no,pic.pic_name
				,useenergy
              from GAMEHistory 
			left join PICDATA pic on pic.ser_no = pic_id
			left join (
				select pic_id,sum(useenergy)-count(pic_id) as useenergy from (
        select LOGIN_ID,pic_id,count(picState) as useenergy
                 from GAMEHistory where (picState='N' or picState='Y' )
			and login_id = @login_id
                group by pic_id,LOGIN_ID
        ) G 
        group by pic_id,LOGIN_ID
			) sg on sg.pic_id = GAMEHistory.pic_id
			where login_id = @login_id
            and (picstate = 'Y' or picstate = 'F')
            order by gametime desc
        ";
        dt = SqlHelper.GetDataTable("PuzzleConnString", sql, 
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", login_id));
        Pager = dt.Paging(pageNumber, pageSize);
        rpList.DataSource = Pager;
        rpList.DataBind();
        SetControl();
        string urlTemp = kmwebsysSite + "/kmactivity/kmwebpuzzle/UserGameLogExport.aspx?querymember=" + HttpUtility.UrlEncode(login_id);
        linkExport.NavigateUrl = urlTemp;
        UserName.Text = login_id;
        sql = @"
            select account.*,JJ.useenergy from account 
            left join (
            select LOGIN_ID,sum(useenergy)-count(pic_id) as useenergy from (
                    select LOGIN_ID,pic_id,count(picState) as useenergy

                             from GAMEHistory where  ( picState='N' or picState='Y' ) 
				            and LOGIN_ID = @login_id
                            group by LOGIN_ID,pic_id
                    ) G 
                    group by LOGIN_ID
            ) JJ on JJ.LOGIN_ID = account.login_id
            where  account.login_id = @login_id
        ";
        dt = SqlHelper.GetDataTable("PuzzleConnString", sql, 
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", login_id));
        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            NickName.Text = GetUserName(dr["nickname"].ToString(), dr["realname"].ToString());
            UserMail.Text = dr["email"].ToString();
            LblEnergy.Text = dr["energy"].ToString();
            LblUseEnergy.Text = dr["useenergy"].ToString();
            LblAllEnergy.Text = dr["getenergy"].ToString();
        }
    }

    private string GetUserName(string nickname,string realname)
    {
        string str = "";
        if (nickname != null && !string.IsNullOrEmpty(nickname.ToString()))
        {
            return nickname.ToString();
        }
        else
        {
            str = realname.ToString();
            return str;
        }
    }

    protected void preLinkAct(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex--;
        DisplayData(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
    protected void nextLinkAct(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex++;
        DisplayData(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
    protected void ChangePageNumber(object sender, EventArgs e)
    {
        DisplayData(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
    }
    protected void ChangePageSize(object sender, EventArgs e)
    {
        DisplayData(0, Convert.ToInt32(PageSizeDDL.SelectedValue));
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

    protected string GetStates(object difficult)
    {
        if (difficult.ToString() == "Y")
        {
            return "完成";
        }
        else
        {
            return "放棄";
        }
    }

    protected string GetDifficult(object difficult)
    {
        if (difficult.ToString() == "H")
        {
            return "完整";
        }
        else
        {
            return "簡單";
        }
    }
    

    protected string GetScore(object difficult,object picstates)
    {
        if (picstates.ToString() == "Y")
        {
            if (difficult.ToString() == "E")
            {
                return "1";
            }
            else
            {
                return "2";
            }
        }
        else
        {
            return "0";
        }
    }
}