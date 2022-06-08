using System;
using System.Collections;
using System.Web.UI;
using System.Web.UI.WebControls;
using PlantLog.Core;
using PlantLog.Core.Domain;
using PlantLog.Core.Service;
using System.Data.SqlClient;

public partial class uploadentry : Page
{
    private IPlantLogService plantLogService;
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
            if (bool.Parse((string)ViewState["isBeforeUpload"]))
            {
                Response.Write("<script language='javascript'>alert('活動尚未開始，敬請期待！');location.href='"+WebUtility.GetAppSetting("WWWUrl")+"';</script>");
            }

            if (bool.Parse((string)ViewState["isAfterUpload"]))
            {
                Response.Write("<script language='javascript'>alert('本活動已截止！請點選參賽作品，開始投票！');location.href='vote.aspx';</script>");
            }

            if (plantLogService == null)
            {
                plantLogService = Utility.ApplicationContext["PlantLogService"] as IPlantLogService;
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

        if (fromBackEnd)
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
                string act = Request.Form["Action"];
                if (act.Contains("DeleteEntry"))
                {
                    DeleteEntry(act.Split('|')[1]);
                }

                if (act.Contains("EditEntry"))
                {
                    EditEntry(act.Split('|')[1]);
                }

                if (act == "UpdateOwnerInfo")
                {
                    UpdateOwnerInfo();
                }

                if (act == "NewLog")
                {
                    NewLog();
                }
            }

            BindData();
        }
    }

    protected void NewLog()
    {
        Response.Redirect("singleentry.aspx?ownerid=" + OwnerId);
    }

    private void BindData()
    {
        entryRepeater.DataSource = Source;
        entryRepeater.DataBind();
    }

    private void EditEntry(string entryId)
    {
        Response.Redirect("singleentry.aspx?ownerid=" + OwnerId + "&entryid=" + entryId);
    }

    private void DeleteEntry(string entryId)
    {
        plantLogService.DeleteEntry(entryId);
        Source = plantLogService.GetEntryByOwner(OwnerId);
    }

    protected void entryRepeater_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            UserControls_EntryControl entryControl = e.Item.FindControl("EntryControl1") as UserControls_EntryControl;
            entryControl.Source = e.Item.DataItem as Entry;
            entryControl.IsEditable = true;
            entryControl.IsAdmin = false;
            entryControl.IsVote = !(bool.Parse((string)ViewState["isAfterVote"]) || bool.Parse((string)ViewState["isBeforeVote"]));
            entryControl.IsOwner = true;
        }
    }

    protected void UpdateOwnerInfo()
    {
        Response.Redirect("editownerinfo.aspx?ownerid=" + OwnerId);
    }

    private void SetOwner()
    {
        if (OwnerId == null)
        {
            OwnerId = Session["memID"].ToString();

            if (WebUtility.IsAdmin(OwnerId))
            {
                Response.Redirect("Default.aspx");
            }
            else
            {
                Owner owner = plantLogService.GetOwner(OwnerId);
                if (owner == null)
                {
                    owner = new Owner();
                    owner.OwnerId = OwnerId;
                    owner.DisplayName = Session["memName"].ToString();
                    owner.Email = GetMail(OwnerId);
                    if (Session["memNickName"] != null)
                    {
                        owner.Nickname = Session["memNickName"].ToString();
                    }
                    owner.Topic = "尚未輸入作品名稱";
                    owner.Description = "尚未輸入作品介紹";
                    owner.CreatorId = "System";
                    owner.CreateDateTime = DateTime.Now;
                    owner.ModifierId = "System";
                    owner.ModifyDateTime = DateTime.Now;

                    UserInfo1.thisOwner = plantLogService.CreateOwner(owner);
                }
                else
                {
                    UserInfo1.thisOwner = owner;
                }
            }
        }
    }

    private void SetListSource()
    {
        if (Source == null)
        {
            Source = plantLogService.GetEntryByOwner(OwnerId);
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
}
