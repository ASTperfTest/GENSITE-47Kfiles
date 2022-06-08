using System;
using System.Collections;
using System.Web.UI;
using System.Web.UI.WebControls;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;
using Gardening.Core.Persistence;
using System.Data.SqlClient;
using System.Data;

public partial class uploadentry : Page
{
    private IGardeningService gardeningService;
    private bool fromBackEnd = false;
    private string OwnerId
    {
        get
        {
            if (ViewState["OwnerId"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["OwnerId"] as string;
            }
        }
        set
        {
            ViewState["OwnerId"] = value;
        }
    }

    private string topicId
    {
        get
        {
            if (ViewState["topicId"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["topicId"] as string;
            }
        }
        set
        {
            ViewState["topicId"] = value;
        }
    }

    private IList Source
    {
        get
        {
            if (ViewState["OwnerEntry"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["OwnerEntry"] as IList;
            }
        }
        set
        {
            ViewState["OwnerEntry"] = value;
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

            SetOwner();

            SetListSource();
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

		if (Session["FromBackEnd"] != null && Convert.ToBoolean(Session["FromBackEnd"]) == true)
        {
            this.MasterPageFile = WebUtility.managementMasterPageFile;
        }
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        if (bool.Parse((string)ViewState["hasLogin"]))
        {
            if (Request.Form["Action"] != null)
            {
                topicId = Request.QueryString["topicId"];
                Topic thisTopic = gardeningService.GetTopic(topicId);
                string act = Request.Form["Action"];
                if (act.Contains("DeleteEntry"))
                {
                    DeleteEntry(act.Split('|')[1], thisTopic.TopicId);
                }

                if (act.Contains("EditEntry"))
                {
                    EditEntry(act.Split('|')[1], thisTopic.TopicId);
                }

                if (act == "UpdateTopic")
                {
                    UpdateTopic(thisTopic.TopicId);
                }

                if (act == "NewLog")
                {
                    NewLog(thisTopic.TopicId);
                }

                if (act == "DeleteTopic")
                {
                    DeleteTopic(thisTopic.TopicId);
                }

				if (act.Contains("ApproveEntry"))
                {
                    ApproveEntry(act.Split('|')[1], thisTopic.TopicId);
                }

                if (act.Contains("HideEntry"))
                {
                    HideEntry(act.Split('|')[1], thisTopic.TopicId);
                }
				
				if (act == "HideOwnerInfo")
                {
                    HideOwnerInfo();
                }

                if (act == "ApproveOwnerInfo")
                {
                    ApproveOwnerInfo();
                }
				
				if (act == "Cancel")
                {
                    RedirectToTopicList();
                }
            }

            BindData();
        }
    }

    protected void NewLog(string id)
    {
        Response.Redirect("singleentry.aspx?ownerid=" + OwnerId + "&topicid=" + id);
    }

    private void BindData()
    {
        entryRepeater.DataSource = Source;
        entryRepeater.DataBind();
    }

    private void EditEntry(string entryId, string id)
    {
        Response.Redirect("singleentry.aspx?ownerid=" + OwnerId + "&entryid=" + entryId + "&topicid=" + id);
    }

    private void DeleteEntry(string entryId,string id)
    {
		//delete kpi data
		Entry thisEntry = gardeningService.GetEntry(entryId);
        string strCreateDate = thisEntry.CreateDateTime.ToShortDateString();
		//Response.Write("<script language=javascript>alert('" + strCreateDate +"!')</script>");
		
		using (SqlConnection tSqlConn = new SqlConnection())
        {
            string webConfigConnectionString1 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

            tSqlConn.ConnectionString = webConfigConnectionString1;
            tSqlConn.Open();

            string sql = @"SELECT memberId FROM MemberGradeShare "+  
                                            "WHERE (memberId =@memberId)  " +
                                            "AND (CONVERT(nvarchar, shareDate, 111) = @creatDate)";
            SqlCommand Cmd = new SqlCommand(sql, tSqlConn);
            Cmd.Parameters.AddWithValue("memberId", Session["memID"]);
            Cmd.Parameters.AddWithValue("creatDate", strCreateDate);

            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommandBuilder cb = new SqlCommandBuilder();
            DataTable dt = new DataTable();

            da.SelectCommand = Cmd;
            cb.DataAdapter = da; // 指定DataAdapter
            da.Fill(dt);

            if ( dt != null && dt.Rows.Count >= 1)
            {
				Cmd.CommandText= @"update MemberGradeShare " +
                                "set sharegardening = case when sharegardening >=1 then sharegardening-1 else 0 end " +
                                "WHERE memberId=@memberId and (CONVERT(nvarchar, shareDate, 111) = @creatDate)";	
                
                Cmd.ExecuteNonQuery();
			}
            else
            { 
                // exception
            }
        }       
				
		gardeningService.DeleteEntry(entryId);		
        Response.Redirect("uploadentry.aspx?topicid="+id);
    
    }
	
	private void ApproveEntry(string entryId,string topicId)
    {
        Entry e = gardeningService.GetEntry(entryId);
        e.IsApprove = true;
        gardeningService.UpdateEntry(e);
        Source = gardeningService.GetEntriesByTopic(topicId);
        BindData();
    }

    private void HideEntry(string entryId,string topicId)
    {
        Entry e = gardeningService.GetEntry(entryId);
        e.IsApprove = false;
        gardeningService.UpdateEntry(e);
        Source = gardeningService.GetEntriesByTopic(topicId);
        BindData();
    }
	
	    protected void ApproveOwnerInfo()
    {
        Topic o = gardeningService.GetTopic(topicId);
        o.IsApprove = true;
        gardeningService.UpdateTopic(o);
        UserInfo1.thisTopic = o;
        trApproveUserInfo.Visible = false;
        trHideUserInfo.Visible = true;
    }

    protected void HideOwnerInfo()
    {
        Topic o = gardeningService.GetTopic(topicId);
        o.IsApprove = false;
        gardeningService.UpdateTopic(o);
        UserInfo1.thisTopic = o;
        trApproveUserInfo.Visible = true;
        trHideUserInfo.Visible = false;
    }
	
    protected void entryRepeater_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            UserControls_EntryControl entryControl = e.Item.FindControl("EntryControl1") as UserControls_EntryControl;
            entryControl.Source = e.Item.DataItem as Entry;
            entryControl.IsEditable = true;
			if (Session["FromBackEnd"] != null && Convert.ToBoolean(Session["FromBackEnd"]) == true)
			{
				entryControl.IsAdmin = true;
			}
			else
			{
				entryControl.IsAdmin = false;
			}
            entryControl.IsOwner = true;
        }
    }

    protected void UpdateTopic(string topicId)
    {
        Response.Redirect("editownerinfo.aspx?ownerid=" + OwnerId + "&topicid=" + topicId);
    }

    protected void DeleteTopic(string topicId)
    {
        if (topicId != null)
        {
            if (gardeningService.GetTopic(topicId) != null)
            {
				//取出TOPIC底下所有的Entry資料 todo- 把kip相關的資料update或清空
				IList entryList = gardeningService.GetEntriesByTopic(topicId);
				ArrayList entryCreateDateList = new ArrayList();
				string entryCreateDate = "";
				foreach (Entry item in entryList)
				{
					entryCreateDate = item.CreateDateTime.ToShortDateString();
					entryCreateDateList.Add(entryCreateDate); 
				}

				//kip update
				string webConfigConnectionString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;
				foreach (string item in entryCreateDateList)
				{
					using (SqlConnection tSqlConn = new SqlConnection())
					{
						tSqlConn.ConnectionString = webConfigConnectionString;
						tSqlConn.Open();
						
						string sql = @"update MemberGradeShare " +
										"set sharegardening = case when sharegardening >=1 then sharegardening-1 else 0 end " +
										"WHERE memberId=@memberId and (CONVERT(nvarchar, shareDate, 111) = @creatDate)";	
						SqlCommand Cmd = new SqlCommand(sql, tSqlConn);
						Cmd.Parameters.AddWithValue("memberId", Session["memID"]);
						Cmd.Parameters.AddWithValue("creatDate", item);
						
						Cmd.ExecuteNonQuery();											
					}
				}
				
                //刪除該筆topic
				gardeningService.DeleteTopic(topicId);
            }
            else
            {
                Response.Write("<script language=javascript>alert('欲刪除的資料不存在!')</script>");
            }
        }
		RedirectToTopicList();
    }

    private void SetOwner()
    {
        topicId = Request.QueryString["topicId"];
        if (OwnerId == null)
        {
            OwnerId = Session["memID"].ToString();
            if (topicId != null)
            {
                Topic thisTopic = gardeningService.GetTopic(topicId);
				if (!WebUtility.IsAdmin(Session["memID"].ToString()))
				{
					trApproveUserInfo.Visible = false;
					trHideUserInfo.Visible = false;
				}
				else
				{
					if (thisTopic.IsApprove)
					{
						trApproveUserInfo.Visible = false;
					}
					else
					{
						trHideUserInfo.Visible = false;
					}
				}	
                UserInfo1.thisTopic = thisTopic;
            }
        }
    }

    private void SetListSource()
    {
        if (Source == null)
        {
            Source = gardeningService.GetEntriesByTopic(topicId);
        }
    }

    private void SetViewState()
    {
        if (ViewState["hasLogin"] == null)
        {
            ViewState["hasLogin"] = WebUtility.CheckLogin(Session["memID"]).ToString();
        }

        DateTime now = DateTime.Now;
        DateTime uploadFromDate = DateTime.Parse(WebUtility.GetAppSetting("UploadFromDate"));
        DateTime uploadToDate = DateTime.Parse(WebUtility.GetAppSetting("UploadToDate"));
    }

    private string GetMail(string ownerId)
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
	
	private void RedirectToTopicList()
	{
		if (Session["FromBackEnd"] != null && Convert.ToBoolean(Session["FromBackEnd"]) == true)
        {
            Response.Redirect("topiclist.aspx");
        }
		else
		{
			Response.Redirect("mytopiclist.aspx");
		}	
	}
}
