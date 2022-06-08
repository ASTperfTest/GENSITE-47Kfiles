using System;
using System.Collections;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using PlantLog.Core;
using PlantLog.Core.Domain;
using PlantLog.Core.Service;
using System.Web;
using System.Text;
using System.Data.SqlClient;

public partial class vote : Page
{
    private IPlantLogService plantLogService;
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
            if (plantLogService == null)
            {
                plantLogService = Utility.ApplicationContext["PlantLogService"] as IPlantLogService;
            }

            trExport.Visible = false;

            if (WebUtility.IsAdmin(Session["memID"].ToString()))
            {
                isAdmin = true;

                if (!bool.Parse((string)ViewState["isBeforeVote"]))
                {
                    trExport.Visible = true;
                }
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
                if (Request.Form["Action"].Contains("CheckRecord"))
                {
                    CheckVoteRecord(Request.Form["Action"].Split('|')[1]);
                }

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

	GridView1.Columns[4].Visible = false;
        GridView1.Columns[3].Visible = false;
        GridView1.Columns[2].Visible = false;
        GridView1.Columns[1].Visible = false;

        if (Source.Rows.Count > 0)
        {
            SetGridBottomPagerInfo();
            GridView1.BottomPagerRow.Visible = true;

            DateTime voteFromDate = DateTime.Parse(WebUtility.GetAppSetting("VoteFromDate"));
            DateTime voteToDate = DateTime.Parse(WebUtility.GetAppSetting("VoteToDate"));
            if (!(bool.Parse((string)ViewState["isBeforeVote"]) || bool.Parse((string)ViewState["isAfterVote"])))
            {
                GridView1.Columns[3].Visible = true;
		GridView1.Columns[2].Visible = true;
                GridView1.Columns[1].Visible = true;
            }

            if (isAdmin)
            {
                GridView1.Columns[4].Visible = true;
            }
        }
    }

    private void VoteUser(string userId)
    {
        string voterId = Session["memID"].ToString();
        string msg = DoVote(voterId, userId, CheckVoteState(voterId, userId));

        if (msg.Length > 0)
        {
            WebUtility.WindowAlert(this.Page, msg);
        }

        SetSource();
    }

    private string CheckVoteState(string voterId, string userId)
    {
        IList allVotes = plantLogService.GetVoteRecordByVoter(voterId);
        int voteLimit;
        if (!int.TryParse(WebUtility.GetAppSetting("voteLimit"), out voteLimit))
        {
            voteLimit = 3;
        }
		
		if(DuplicateEmail(voterId))
		{
			return "Voted";
		}

        if (allVotes == null || allVotes.Count == 0)
        {
            return "Non";
        }
        else
        {
            if (allVotes.Count >= voteLimit)
            {
                return "Voted";
            }
            else
            {
                foreach (VoteRecord vr in allVotes)
                {
                    if (vr.UserId == userId)
                    {
                        return "Voted";
                    }
                }
                return "Non";
            }
        }
    }
	
	private bool DuplicateEmail(string voterId)
	{
			string voterEmail = GetEmail(voterId);
			DataTable dt = new DataTable();
			string cmdText =" SELECT 1 FROM" + 
			" ( " +
			" SELECT M.EMAIL,COUNT(*) AS CNT FROM PLANT_LOG..VOTE_RECORD V" +
			" LEFT JOIN MEMBER M" +
			" ON V.VOTER_ID=M.ACCOUNT" +
			" GROUP BY M.EMAIL" +
			" ) AS A" +
			" WHERE A.EMAIL LIKE '%" + voterEmail + "%' AND CNT >= 3";
			SqlCommand sc = new SqlCommand(cmdText);
            sc.Connection = new SqlConnection(WebUtility.GetAppSetting("COAConnectionString"));
            try
            {
                sc.Connection.Open();
                SqlDataAdapter sa = new SqlDataAdapter(sc);
                sa.Fill(dt);
            }
            catch (Exception err)
            {
                throw err;
            }
            finally
            {
                sc.Connection.Close();
            }
			return (dt.Rows.Count==0)?false:true;
	}

    private string DoVote(string voterId, string userId, string voteState)
    {
        string userIP = Request.UserHostAddress;
        string msg = "";

        switch (voteState)
        {
            case "Non":
                VoteRecord voteRecord = new VoteRecord();
                voteRecord.IP = userIP;
                voteRecord.VoteDate = DateTime.Now;
                voteRecord.UserId = userId;
                voteRecord.VoterId = voterId;

                plantLogService.CreateVoteRecord(voteRecord);
                msg = WebUtility.GetAppSetting("Thanks");
                break;
            case "Voted":
                msg = WebUtility.GetAppSetting("Warning");
                break;
            default:
                break;
        }

        return msg;
    }

    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            Owner thisOwner = plantLogService.GetOwner(((DataRowView)e.Row.DataItem)["ID"].ToString());

            e.Row.Cells[1].Text = thisOwner.VoteRecords.Count.ToString();
	    IList en = plantLogService.GetEntryByOwner(((DataRowView)e.Row.DataItem)["ID"].ToString());
	    e.Row.Cells[2].Text = en.Count.ToString();
	
            UserControls_UserInfo user = e.Row.FindControl("userInfo1") as UserControls_UserInfo;
            user.thisOwner = thisOwner;

            Image img = user.FindControl("AvatarImage") as Image;
            if (Session["memID"].ToString() == thisOwner.OwnerId)
            {
                img.Attributes.Add("OnClick",
                    "javascript:window.location='uploadentry.aspx';");
            }
            else
            {
                img.Attributes.Add("OnClick",
                    "javascript:window.location='entrylist.aspx?ownerid=" + thisOwner.OwnerId + "';");
            }
            img.AlternateText = thisOwner.DisplayName;
            img.ToolTip = "點選檢視完整記錄";
        }
    }

    protected void CheckVoteRecord(string ownerId)
    {
        IList records = plantLogService.GetVoteRecordByUser(ownerId);

        DataTable result = new DataTable();
        result.Columns.Add(new DataColumn("VoterId"));
	result.Columns.Add(new DataColumn("VoterName"));
        result.Columns.Add(new DataColumn("VoteDate"));
	result.Columns.Add(new DataColumn("IP"));
		
        foreach (VoteRecord vr in records)
        {
                DataRow dr = result.NewRow();
				dr["VoterId"] = vr.VoterId;
                dr["VoterName"] = GetDisplayName(vr.VoterId);
		dr["IP"] = vr.IP;
                dr["VoteDate"] = (vr.VoteDate.Year - 1911).ToString() + "/" +
                    vr.VoteDate.Month.ToString() + "/" + vr.VoteDate.Day.ToString();

                result.Rows.Add(dr);
        }

        VoteRecordList1.OwnerId = ownerId;
        VoteRecordList1.Source = result;
        VoteRecordList1.Visible = true;
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

		if (GridView1.PageCount<=1){
			pageTextBox.Visible=false;
			LinkButton lb1=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton1") as LinkButton;
			LinkButton lb2=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton2") as LinkButton;
			LinkButton lb3=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton3") as LinkButton;
			LinkButton lb4=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton4") as LinkButton;
			lb1.Visible=false;
			lb2.Visible=false;
			lb3.Visible=false;
			lb4.Visible=false;
		}else if(GridView1.PageCount==Convert.ToInt32(pageTextBox.Text)){
			LinkButton lb3=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton3") as LinkButton;
			LinkButton lb4=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton4") as LinkButton;
			lb3.Visible=false;
			lb4.Visible=false;
		}else if(Convert.ToInt32(pageTextBox.Text)==1){
			LinkButton lb1=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton1") as LinkButton;
			LinkButton lb2=this.GridView1.BottomPagerRow.Cells[0].FindControl("LinkButton2") as LinkButton;
			lb1.Visible=false;
			lb2.Visible=false;
		}
		
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
        IList result = plantLogService.GetAllOwner();

        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add(new DataColumn("ID",typeof(string)));
        dtTemp.Columns.Add(new DataColumn("Name",typeof(string)));
        dtTemp.Columns.Add(new DataColumn("Nickname",typeof(string)));
        dtTemp.Columns.Add(new DataColumn("email",typeof(string)));
        dtTemp.Columns.Add(new DataColumn("VoteCount",typeof(int)));
        dtTemp.Columns.Add(new DataColumn("Topic",typeof(string)));
        dtTemp.Columns.Add(new DataColumn("EntryCount",typeof(int)));
        dtTemp.Columns.Add(new DataColumn("LastModifyDateTime",typeof(DateTime)));

        foreach (Owner temp in result)
        {
            if ((isAdmin || temp.IsApprove || temp.OwnerId == Session["memID"].ToString()) && (temp.Topic != "尚未輸入作品名稱"))
            {
                DataRow dr = dtTemp.NewRow();

                dr["ID"] = temp.OwnerId;
                dr["Name"] = temp.DisplayName;
                if (temp.Nickname == null)
                {
                    dr["Nickname"] = string.Empty;
                }
                else
                {
                    dr["Nickname"] = temp.Nickname;
                }
                if (temp.Email == null)
                {
                    dr["email"] = "";
                }
                else
                {
                    dr["email"] = temp.Email;
                }
                dr["VoteCount"] = temp.VoteRecords.Count;
                dr["Topic"] = temp.Topic;

                IList en = plantLogService.GetEntryByOwner(temp.OwnerId);
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

                dtTemp.Rows.Add(dr);
            }
        }

        Source = dtTemp;
    }
	
	private string GetDisplayName(string ownerId)
    {
        string strSQL = "SELECT realname FROM Member WHERE account = '" + ownerId + "'";
        string connectionString = WebUtility.GetAppSetting("COAConnectionString");
        string result = null;

        using (SqlConnection cn = new SqlConnection(connectionString))
        {
            SqlCommand objsqlcmd = new SqlCommand(strSQL, cn);
            cn.Open();

            SqlDataReader reader = objsqlcmd.ExecuteReader();

            // Call Read before accessing data.
            while (reader.Read())
            {
                result = HttpUtility.HtmlDecode(reader[0].ToString());
            }

            // Call Close when done reading.
            reader.Close();
        }

        return result;
    }
	
	private string GetEmail(string ownerId)
    {
        string strSQL = "SELECT email FROM Member WHERE account = '" + ownerId + "'";
        string connectionString = WebUtility.GetAppSetting("COAConnectionString");
        string result = null;

        using (SqlConnection cn = new SqlConnection(connectionString))
        {
            SqlCommand objsqlcmd = new SqlCommand(strSQL, cn);
            cn.Open();

            SqlDataReader reader = objsqlcmd.ExecuteReader();

            // Call Read before accessing data.
            while (reader.Read())
            {
                result = reader[0].ToString();
            }

            // Call Close when done reading.
            reader.Close();
        }

        return result;
    }
}
