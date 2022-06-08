using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using GSS.Vitals.COA.Data;

public partial class Century_Story_List : System.Web.UI.Page
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
        string iCTUnit = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnit"].ToString();
        string sqlQueryScript;
        sqlQueryScript = @"SELECT CuDTGeneric.*, InfoUser.UserName As UserName
                           FROM CuDTGeneric LEFT JOIN InfoUser ON CuDTGeneric.iEditor = InfoUser.UserID
                           WHERE iCTUnit = @iCTUnit AND fCTUPublic = 'Y' 
                           ORDER BY xPostDate DESC";
        dt = SqlHelper.GetDataTable("ODBCDSN", sqlQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN","@iCTUnit","@iCTUnit",iCTUnit));

       
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
        //Response.Write(Pager.CurrentPageIndex);
        PageNumberDDL.SelectedIndex = Pager.CurrentPageIndex;

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

    protected void rptList_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            string strURL = System.Web.Configuration.WebConfigurationManager.AppSettings["WWWUrl"];
            ((Image)e.Item.FindControl("imgArticle")).ImageUrl = strURL + "/Public/Data/" +((Image)e.Item.FindControl("imgArticle")).ImageUrl;
            ((HyperLink)e.Item.FindControl("HyperLink1")).NavigateUrl = strURL + "/Century/Story_Detail.aspx?xItem=" + ((Label)e.Item.FindControl("labICUItem")).Text;
            ((Label)e.Item.FindControl("labDate")).Text = Convert.ToDateTime(((Label)e.Item.FindControl("labDate")).Text).ToString("yyyy.M.d");
            if (HtmlRemoval.StripTagsRegex(((Label)e.Item.FindControl("labContent")).Text).Length < 250)
            {
                ((Label)e.Item.FindControl("labContent")).Text = HtmlRemoval.StripTagsRegex(((Label)e.Item.FindControl("labContent")).Text);
            }
            else
            {
                ((Label)e.Item.FindControl("labContent")).Text = HtmlRemoval.StripTagsRegex(((Label)e.Item.FindControl("labContent")).Text).Substring(0,250) + "..";
            }            
        }
    }

    
}