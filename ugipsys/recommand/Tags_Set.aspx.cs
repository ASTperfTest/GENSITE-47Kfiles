using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Transactions;
using System.Data;

public partial class Tags_Set : System.Web.UI.Page
{
    private PagedDataSource Pager;
    private DataTable dt;

    string Script = "<script>alert('連線逾時或尚未登入，請登入會員');window.close();</script>";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()))
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
        }
    }

    // 新增
    protected void btnAdd_Click(object sender, EventArgs e)
    {
        try
        {
            string sqlInsertScript;
            string strQueryScript = @"SELECT COUNT(*) FROM TAGs WHERE tagName = @tagName";

            if (!string.IsNullOrEmpty(txtSearch.Text) && Session["userID"] != null 
                                && !string.IsNullOrEmpty(Session["userID"].ToString()))
            {
                int intCount = (int)SqlHelper.ReturnScalar("ConnString", strQueryScript,
                    DbProviderFactories.CreateParameter("ConnString", "@tagName", "@tagName", txtSearch.Text));
                if (intCount <= 0)
                {
                    sqlInsertScript = @"INSERT INTO TAGs (tagName, tagCreator, createdDate) VALUES (@tagName, @tagCreator, @createdDate)";
                    rptList.DataSource = SqlHelper.GetDataTable("ConnString", sqlInsertScript,
                        DbProviderFactories.CreateParameter("ConnString", "@tagName", "@tagName", txtSearch.Text),
                        DbProviderFactories.CreateParameter("ConnString", "@tagCreator", "@tagCreator", Session["userID"].ToString()),
                        DbProviderFactories.CreateParameter("ConnString", "@createdDate", "@createdDate", DateTime.Now.ToString("yyyy/MM/dd")));
                }
                else
                {
                    Response.Write("<script language='javascript'>alert('標籤不可為重覆!');location.href('Tags_Set.aspx');</script>");
                }
            }
            else if(string.IsNullOrEmpty(txtSearch.Text))
            {
                Response.Write("<script language='javascript'>alert('標籤不可為空!');location.href('Tags_Set.aspx');</script>");
            }
            else if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()))
            {
                Response.Write("<script language='javascript'>alert('請重新登入!');close();</script>");
            }
            Response.Write("<script language='javascript'>alert('新增完成');location.href('Tags_Set.aspx');</script>");
        }
        catch (Exception)
        {
            //Response.Write("<script language='javascript'>alert('發生錯誤，無法將tag加入至資料表中');location.href('Tags_Set.aspx');</script>");
            throw;
        }
    }

    // 搜尋
    protected void btnSearch_Click(object sender, EventArgs e)
    {
        // Page_Load(null,null);
    }

    // 刪除
    protected void btnDelete_Click(object sender, EventArgs e)
    {
        using (TransactionScope Scope = new TransactionScope())
        {
            try
            {
                for (int i = 0; i < rptList.Items.Count; i++)
                {
                    if (((CheckBox)rptList.Items[i].FindControl("checkbox1")).Checked)
                    {
                        string labtagID = ((Label)rptList.Items[i].FindControl("labtagID")).Text;
                        string sqlDeleteScript = @"DELETE FROM     TAGs 
                                           WHERE        tagID = @tagID";
                        SqlHelper.ExecuteNonQuery("ConnString", sqlDeleteScript,
                            DbProviderFactories.CreateParameter("ConnString", "@tagID", "@tagID", labtagID));
                    }
                }
                myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), 10);
                Scope.Complete();
            }
            catch (Exception)
            {
                Response.Write("<script language=" + "javascript>" + "alert('刪除的選項包含已被使用的Tag，故無法動作');</script>");
                //throw;
            }
        }
    }

    protected void myDBinit(int intPageNumber, int intPageSize)
    {
        string sqlQueryScript;
        if (!string.IsNullOrEmpty(txtSearch.Text))
        {
            sqlQueryScript = @"SELECT       TAGs.tagID, TAGs.tagName, tmpTable.intCount 
                               FROM         TAGs LEFT JOIN
                                    (SELECT     tagID, COUNT(*) AS intCount
                                     FROM       RecommandContent2TAGs
                                     GROUP BY   tagID)  AS  tmpTable
                               ON  TAGs.TagID = tmpTable.TagID
                               WHERE TAGs.tagName LIKE '%' + @tagName + '%' ORDER BY TAGs.tagID";
            dt = SqlHelper.GetDataTable("ConnString", sqlQueryScript,
                DbProviderFactories.CreateParameter("ConnString", "@tagName", "@tagName", txtSearch.Text));
        }
        else
        {
            sqlQueryScript = @"SELECT       TAGs.tagID, TAGs.tagName, tmpTable.intCount 
                               FROM         TAGs LEFT JOIN
                                    (SELECT     tagID, COUNT(*) AS intCount
                                     FROM       RecommandContent2TAGs
                                     GROUP BY   tagID)  AS  tmpTable
                               ON  TAGs.TagID = tmpTable.TagID";
            dt = SqlHelper.GetDataTable("ConnString", sqlQueryScript);
        }
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
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), 10);
    }

    protected void PreviousLink_Click(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex--;
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), 10);
    }

    protected void NextLink_Click(object sender, EventArgs e)
    {
        PageNumberDDL.SelectedIndex++;
        myDBinit(Convert.ToInt32(PageNumberDDL.SelectedValue), 10);
    }

    // 設定checkbox是否Enable
    protected void rptList_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            if (!string.IsNullOrEmpty(((Label)e.Item.FindControl("labCount")).Text))
            {
                ((CheckBox)e.Item.FindControl("checkbox1")).Enabled = false;
            }
        }
    }
}