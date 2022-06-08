using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Text;
using System.Data;

/// <summary>
/// Added By Leo    2011-06-27      我的好文推薦文章列表
/// </summary>
public partial class Recommand_Mylist : System.Web.UI.Page
{
    private PagedDataSource Pager;
    private DataTable dt;

    string MemberID;
    string Script = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>";

    // Tab_Init
    protected void Page_Init(object sender, EventArgs e)
    {
        TabText2.Text = WebUtility.GetMyAreaLinks(MyAreaLink.recommand_Mylist);
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null || string.IsNullOrEmpty(Session["memID"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
        else
        {
            if (!IsPostBack)
            {
                int PageNumber = 0;
                int PageSize = 10;

                myDBinit(PageNumber, PageSize);
            }
            else
            {
                myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), Convert.ToInt32(PageSizeDDL.SelectedValue));
            }
        }
    }

    protected void myDBinit(int PageNumber, int PageSize)
    {
        MemberID = Session["memID"].ToString();

        string sqlQueryScript = @"SELECT * FROM RecommandContent WHERE iEditor = @iEditor ORDER BY aEditDate desc";
        dt = SqlHelper.GetDataTable("ConnString", sqlQueryScript,
            DbProviderFactories.CreateParameter("ConnString", "@iEditor", "@iEditor", MemberID));
        Pager = dt.Paging(PageNumber, PageSize);
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




    protected void rptList_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            // 推薦日期
            DateTime dt = Convert.ToDateTime(((Label)e.Item.FindControl("labDate")).Text);
            ((Label)e.Item.FindControl("labDate")).Text = dt.ToString("yyyy/MM/dd");

            // 通過審查
            switch (((Label)e.Item.FindControl("labExam")).Text)
            {
                case "y":
                    ((Label)e.Item.FindControl("labExam")).ForeColor = System.Drawing.Color.Blue;
                    ((Label)e.Item.FindControl("labExam")).Text = "已通過";
                    break;
                case "n":
                    ((Label)e.Item.FindControl("labExam")).ForeColor = System.Drawing.Color.Red;
                    ((Label)e.Item.FindControl("labExam")).Text = "未通過";
                    break;
                default:
                    ((Label)e.Item.FindControl("labExam")).Text = "尚未審查";
                    break;
            }

        }
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

}