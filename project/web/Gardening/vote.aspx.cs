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

public partial class vote : Page
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
            if (Request.ServerVariables["HTTP_REFERER"] != null && (Request.ServerVariables["HTTP_REFERER"].ToLower().IndexOf(WebUtility.GetAppSetting("kmwebsysSite").ToString()) > -1))
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
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        SetViewState();
        if (bool.Parse((string)ViewState["hasLogin"]))
        {
            if (gardeningService == null)
            {
                gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
            }

            if (WebUtility.IsAdmin(Session["memID"].ToString()))
            {
                isAdmin = true;
            }
            else
            {
                isAdmin = false;
            }

            VoteRecordList1.Visible = false;
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "MyScript",
                "alert('請先登入會員');setHref('" + WebUtility.GetAppSetting("RedirectPage") + "');", true);
        }
        if (!IsPostBack)
        {
            useManagementMasterPage.Value = fromBackEnd.ToString();
        }
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        if (bool.Parse((string)ViewState["hasLogin"]))
        {
            if (Request.Form["Action"] != null)
            {

                if (Request.Form["Action"].Contains("Vote"))
                {
                    VoteUser(Request.Form["Action"].Split('|')[1]);
                }

                if (Request.Form["Action"].Contains("Export"))
                {
                    Export();
                }
            }

            SetSource();

            BindData();
        }
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

            DateTime voteFromDate = DateTime.Parse(WebUtility.GetAppSetting("VoteFromDate"));
            DateTime voteToDate = DateTime.Parse(WebUtility.GetAppSetting("VoteToDate"));
        }
    }

    private void VoteUser(string userId)
    {
        string voterId = Session["memID"].ToString();

        SetSource();
    }


    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            Topic thisTopic = gardeningService.GetTopic(((DataRowView)e.Row.DataItem)["TopicId"].ToString());

            //e.Row.Cells[1].Text = thisOwner.VoteRecords.Count.ToString();

            UserControls_UserInfo user = e.Row.FindControl("userInfo1") as UserControls_UserInfo;
            user.thisTopic = thisTopic;

            Image img = user.FindControl("AvatarImage") as Image;
            if (Session["memID"].ToString() == thisTopic.Owner.UserId)
            {
                img.Attributes.Add("OnClick",
                    "javascript:window.location='uploadentry.aspx?topicid=" + thisTopic.TopicId + "';");
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


    private void SetViewState()
    {
        if (ViewState["hasLogin"] == null)
        {
            ViewState["hasLogin"] = WebUtility.CheckLogin(Session["memID"]).ToString();
        }

        DateTime now = DateTime.Now;
        DateTime voteFromDate = DateTime.Parse(WebUtility.GetAppSetting("VoteFromDate"));
        DateTime voteToDate = DateTime.Parse(WebUtility.GetAppSetting("VoteToDate"));
        DateTime uploadFromDate = DateTime.Parse(WebUtility.GetAppSetting("UploadFromDate"));
        DateTime uploadToDate = DateTime.Parse(WebUtility.GetAppSetting("UploadToDate"));

        if (ViewState["isBeforeUpload"] == null || ViewState["isAfterUpload"] == null)
        {
            if (DateTime.Compare(now, uploadFromDate) > 0)
            {
                ViewState["isBeforeUpload"] = false.ToString();
            }
            else
            {
                ViewState["isBeforeUpload"] = true.ToString();
            }

            if (DateTime.Compare(now, uploadToDate) > 0)
            {
                ViewState["isAfterUpload"] = true.ToString();
            }
            else
            {
                ViewState["isAfterUpload"] = false.ToString();
            }
        }

        if (ViewState["isBeforeVote"] == null || ViewState["isAfterVote"] == null)
        {
            if (DateTime.Compare(now, voteFromDate) > 0)
            {
                ViewState["isBeforeVote"] = false.ToString();
            }
            else
            {
                ViewState["isBeforeVote"] = true.ToString();
            }

            if (DateTime.Compare(now, voteToDate) > 0)
            {
                ViewState["isAfterVote"] = true.ToString();
            }
            else
            {
                ViewState["isAfterVote"] = false.ToString();
            }
        }
    }

    private void Export()
    {
        if (Source.Rows.Count > 0)
        {
            DataTable dtTemp = new DataTable();
            dtTemp.Columns.Add(new DataColumn("帳號"));
            dtTemp.Columns.Add(new DataColumn("姓名"));
            dtTemp.Columns.Add(new DataColumn("暱稱"));
            dtTemp.Columns.Add(new DataColumn("email"));
            dtTemp.Columns.Add(new DataColumn("票數"));

            foreach (DataRow d in Source.Rows)
            {
                DataRow dr = dtTemp.NewRow();

                dr["帳號"] = d["ID"];
                dr["姓名"] = d["Name"];
                dr["暱稱"] = d["Nickname"];
                dr["email"] = d["email"];
                dr["票數"] = d["VoteCount"];

                dtTemp.Rows.Add(dr);
            }

            ExportToExcel(dtTemp);
        }
    }

    private void ExportToExcel(DataTable table)
    {
        HttpContext context = HttpContext.Current;
        context.Response.Clear();
        context.Response.ContentEncoding = Encoding.GetEncoding("Big5");

        foreach (DataColumn column in table.Columns)
        {
            context.Response.Write(column.ColumnName + ",");
        }

        context.Response.Write(Environment.NewLine);

        foreach (DataRow row in table.Rows)
        {
            for (int i = 0; i < table.Columns.Count; i++)
            {
                context.Response.Write(row[i].ToString().Replace(",", string.Empty) + ",");
            }
            context.Response.Write(Environment.NewLine);
        }

        context.Response.ContentType = "text/csv";
        context.Response.AppendHeader("Content-Disposition", "attachment; filename=VoteRecord.csv");
        context.Response.End();
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
        //BindGridView
        BindData();
    }

    protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        if (e.NewPageIndex >= 0)
        {
            GridView1.PageIndex = e.NewPageIndex;
            //BindGridView
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
        IList result = gardeningService.GetAllTopics();

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
            if ((isAdmin || temp.IsApprove || temp.Owner.UserId == Session["memID"].ToString()) && (temp.Title != "尚未輸入作品名稱"))
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
