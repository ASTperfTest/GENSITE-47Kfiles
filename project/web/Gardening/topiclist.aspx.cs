using System;
using System.Collections;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;
using System.Web;
using System.Text;

public partial class topiclist : Page
{
    private IGardeningService gardeningService;
    private bool isAdmin;
    private bool fromBackEnd = false;
    private DataTable Source
    {
        get
        {
            if (ViewState["AllUser"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["AllUser"] as DataTable;
            }
        }
        set
        {
            ViewState["AllUser"] = value;
        }
    }

    protected override void OnPreInit(EventArgs e)
    {
        if (!IsPostBack)
        {
            if ((Request.ServerVariables["HTTP_REFERER"] != null && (Request.ServerVariables["HTTP_REFERER"].ToLower().IndexOf(WebUtility.GetAppSetting("kmwebsysSite").ToString()) > -1)) || (Session["FromBackEnd"] != null && Convert.ToBoolean(Session["FromBackEnd"])))
            {
                WebUtility.SetManagerSessions();
                ViewState["hasLogin"] = true.ToString();
                fromBackEnd = true;
            }
        }
        else
        {
            string hiddenFromBackEnd = Request.Form["ctl00$cp$useManagementMasterPage"];
            if (hiddenFromBackEnd != null && hiddenFromBackEnd != "" && Convert.ToBoolean(hiddenFromBackEnd))
            {
                fromBackEnd = true;
            }
        }

        if (fromBackEnd)
        {
            this.MasterPageFile = WebUtility.managementMasterPageFile;
			Session["FromBackEnd"]=true.ToString();
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (gardeningService == null)
        {
            gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        }
        if (ViewState["hasLogin"] != null)
        {

            if (WebUtility.IsAdmin(Session["memID"].ToString()))
            {
                isAdmin = true;
            }
        }
        else
        {
            isAdmin = false;
        }
        if (!IsPostBack)
        {
            useManagementMasterPage.Value = fromBackEnd.ToString();
            SetSource();
            BindData();
        }
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
    }

    private void BindData()
    {
        DataView dv = Source.DefaultView;
        dv.Sort = RadioButtonList2.SelectedValue + " " + RadioButtonList1.SelectedValue;

        GridView1.DataSource = dv;
        GridView1.DataBind();


        if (Source.Rows.Count > 0)
        {
            SetGridBottomPagerInfo();
            GridView1.BottomPagerRow.Visible = true;
        }
    }

    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            Topic thisTopic = gardeningService.GetTopic(((DataRowView)e.Row.DataItem)["TopicId"].ToString());

            //e.Row.Cells[1].Text = thisOwner.VoteRecords.Count.ToString();

            UserControls_UserInfo user = e.Row.FindControl("UserInfo1") as UserControls_UserInfo;
            user.thisTopic = thisTopic;

            Image img = user.FindControl("AvatarImage") as Image;
            if (Session["memID"] != null)
            {
				bool isAdmin = WebUtility.IsAdmin(Session["memID"].ToString());
                if ((Session["memID"].ToString() == thisTopic.Owner.UserId) || isAdmin == true )
                {
                    img.Attributes.Add("OnClick",
                        "javascript:window.location='uploadentry.aspx?topicid=" + thisTopic.TopicId + "';");
                }
                else
                {
                    img.Attributes.Add("OnClick",
                        "javascript:window.location='entrylist.aspx?topicid=" + thisTopic.TopicId + "';");
                }
            }
            else
            {
                img.Attributes.Add("OnClick",
                       "javascript:window.location='entrylist.aspx?topicid=" + thisTopic.TopicId + "';");
            }
            img.AlternateText = thisTopic.Owner.DisplayName;
            img.ToolTip = "點選檢視完整記錄";
        }
    }

    #region Paging

    private void SetGridBottomPagerInfo()
    {
        Label pageTextBox = this.GridView1.BottomPagerRow.Cells[0].FindControl("PageIndexTextBox") as Label;
        pageTextBox.Text = Convert.ToString(this.GridView1.PageIndex + 1);

        Label totalTuplesTextBox = this.GridView1.BottomPagerRow.Cells[0].FindControl("TotalTuplesTextBox") as Label;
        totalTuplesTextBox.Text = Convert.ToString(Source.Rows.Count);

        Label totalPagesTextBox = this.GridView1.BottomPagerRow.Cells[0].FindControl("TotalPagesTextBox") as Label;
        totalPagesTextBox.Text = Convert.ToString(this.GridView1.PageCount);

        DropDownList numberOfTuplesDDL =
            this.GridView1.BottomPagerRow.Cells[0].FindControl("NumOfTuplesDropDownList") as DropDownList;

        switch (this.GridView1.PageSize)
        {
            case 5:
                numberOfTuplesDDL.SelectedIndex = 0;
                break;
            case 10:
                numberOfTuplesDDL.SelectedIndex = 1;
                break;
            case 50:
                numberOfTuplesDDL.SelectedIndex = 2;
                break;
            default:
                break;
        }
    }

    protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        DropDownList numberOfTuplesDDL =
            GridView1.BottomPagerRow.Cells[0].FindControl("NumOfTuplesDropDownList") as DropDownList;
        GridView1.PageSize = Convert.ToInt32(numberOfTuplesDDL.SelectedValue);
        int selectedIndex = numberOfTuplesDDL.SelectedIndex;
        GridView1.PageIndex = 0;
        BindData();
    }

    protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        if (e.NewPageIndex >= 0)
        {
            GridView1.PageIndex = e.NewPageIndex;
            BindData();
        }
    }

    #endregion

    protected void Sort_Click(object sender, EventArgs e)
    {
        SetSource();
        BindData();
    }

    private void SetSource()
    {
        string Keyword = HttpUtility.HtmlEncode(WebUtility.GetStringParameter("Keyword", string.Empty));
        IList result;
        if (Keyword == null || Keyword == "")
        {
            result = gardeningService.GetAllTopics();
        }
        else
        {
            result = gardeningService.GetTopicsSearchResult(Keyword,0,999);
        }

        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add(new DataColumn("ID"));
        dtTemp.Columns.Add(new DataColumn("Name"));
        dtTemp.Columns.Add(new DataColumn("Nickname"));
        dtTemp.Columns.Add(new DataColumn("email"));
        dtTemp.Columns.Add(new DataColumn("Topic"));
        dtTemp.Columns.Add(new DataColumn("EntryCount"));
        dtTemp.Columns.Add(new DataColumn("LastModifyDateTime"));
        dtTemp.Columns.Add(new DataColumn("TopicId"));

        foreach (Topic temp in result)
        {
			String memberId=(Session["memID"]==null)?"":Session["memID"].ToString();
			if ((isAdmin || temp.IsApprove || temp.Owner.UserId == memberId) && (temp.Title != "尚未輸入作品名稱"))
            {
                DataRow dr = dtTemp.NewRow();

                dr["ID"] = temp.Owner.UserId;
                dr["Name"] = temp.Owner.DisplayName;
                if (temp.Owner.Nickname == null)
                {
                    dr["Nickname"] = string.Empty;
                }
                else
                {
                    dr["Nickname"] = temp.Owner.Nickname;
                }
                if (temp.Owner.Email == null)
                {
                    dr["email"] = "";
                }
                else
                {
                    dr["email"] = temp.Owner.Email;
                }
                dr["Topic"] = temp.Title;

                IList en = gardeningService.GetEntriesByTopic(temp.TopicId);
                dr["EntryCount"] = en.Count;

                DateTime last = DateTime.MinValue;

                foreach (Entry e in en)
                {
                    if (DateTime.Compare(e.ModifyDateTime, last) > 0)
                    {
                        last = e.ModifyDateTime;
                    }
                }
                dr["LastModifyDateTime"] = last.ToLongDateString();
                dr["TopicId"] = temp.TopicId;

                dtTemp.Rows.Add(dr);
            }
        }

        Source = dtTemp;
    }
}
