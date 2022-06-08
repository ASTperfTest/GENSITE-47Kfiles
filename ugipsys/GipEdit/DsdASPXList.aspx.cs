using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Data;

public partial class DsdASPXList : System.Web.UI.Page
{
    private PagedDataSource Pager;
    private DataTable dt;
    private string iCTUnit, WebURL;

    protected void Page_Init(object sender, EventArgs e)
    {
        myValueInit();
        iCTUnit = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnit"].ToString();
        WebURL = System.Web.Configuration.WebConfigurationManager.AppSettings["WebURL"].ToString();
    }
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            int PageNumber = 0;
            int PageSize = 10;

            myDBinit(PageNumber, PageSize);
        }
    }

    private void myValueInit()
    {
        string strQueryString = @"SELECT u.*, u.ibaseDsd, b.sbaseTableName, r.pvXdmp, catName FROM CatTreeNode AS n 
                                    LEFT JOIN CtUnit AS u ON u.ctUnitId=n.ctUnitId Left Join BaseDsd As b 
                                        ON u.ibaseDsd=b.ibaseDsd LEFT JOIN CatTreeRoot AS r ON r.ctRootId=n.ctRootId
                                  WHERE n.ctNodeId= @ctNodeId";
        using (var reader = SqlHelper.ReturnReader("ConnString", strQueryString,
            DbProviderFactories.CreateParameter("ConnString", "@ctNodeId", "@ctNodeId", Request.QueryString["ctNodeId"].ToString())))
        {
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    Session["ctNodeId"] = Request.QueryString["ctNodeId"].ToString();
                    Session["ItemID"] = Request.QueryString["ItemID"].ToString();
                    Session["ctUnitId"] = reader["ctUnitId"].ToString();
                    Session["ctUnitName"] = reader["ctUnitName"].ToString();
                    Session["catName"] = reader["catName"].ToString();
                    Session["ibaseDsd"] = reader["ibaseDsd"].ToString();
                    Session["fCtUnitOnly"] = reader["fCtUnitOnly"].ToString();
                }
            }
        }
    }

    private void myDBinit(int intPageNumber, int intPageSize)
    {
        string sqlQueryScript = "SELECT * FROM CuDTGeneric WHERE iCTUnit = @iCTUnit AND refId = @refId ";

        if (Request.QueryString["from"] != null)
        {
            string from = Request.QueryString["from"].ToString();
            if (from == "query")
                sqlQueryScript = sqlQueryScript + "AND sTitle LIKE '%" + Request.QueryString["title"].ToString() + "%' ";
            
        }
        
        sqlQueryScript = sqlQueryScript + "ORDER BY xPostDate DESC";
        
        dt = SqlHelper.GetDataTable("ConnString", sqlQueryScript,
            DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnit),
            DbProviderFactories.CreateParameter("ConnString", "@refId", "@refId", Session["CtNodeID"].ToString()));

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
            string strPublic = ((Literal)e.Item.FindControl("litPublic")).Text;
            switch (strPublic)
            {
                case "Y":
                    ((Literal)e.Item.FindControl("litPublic")).Text = "公開";
                    break;
                case "N":
                    ((Literal)e.Item.FindControl("litPublic")).Text = "不公開";
                    break;
                default:
                    break;
            }
            // 預覽
            ((HyperLink)e.Item.FindControl("linkView")).NavigateUrl =
                WebURL + "/Century/Picture_Detail.aspx?ctNodeId=" + Request.QueryString["ctNodeId"].ToString(); 
            // 編修
            ((HyperLink)e.Item.FindControl("linkEdit")).NavigateUrl = "DsdASPXEdit.aspx?iCUItem="
                + ((Literal)e.Item.FindControl("litiCUItem")).Text; 

        }
    }
}