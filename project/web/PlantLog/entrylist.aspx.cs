using System;
using System.Collections;
using System.Web.UI;
using System.Web.UI.WebControls;
using PlantLog.Core;
using PlantLog.Core.Domain;
using PlantLog.Core.Service;

public partial class entrylist : Page
{
    private IPlantLogService plantLogService;
    private bool isAdmin;
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
            SetProperty();

            if (plantLogService == null)
            {
                plantLogService = Utility.ApplicationContext["PlantLogService"] as IPlantLogService;
            }

            SetOwner();

            //防止使用者登入後，利用自行輸入的網址想看未開放的日誌
            if (!UserInfo1.thisOwner.IsApprove)
            {
                if (!isAdmin)
                {
                    Response.Redirect("Default.aspx");
                }
            }

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
                if (act.Contains("ApproveEntry"))
                {
                    ApproveEntry(act.Split('|')[1]);
                }

                if (act.Contains("HideEntry"))
                {
                    HideEntry(act.Split('|')[1]);
                }

                if (act == "HideOwnerInfo")
                {
                    HideOwnerInfo();
                }

                if (act == "ApproveOwnerInfo")
                {
                    ApproveOwnerInfo();
                }
            }

            BindData();
        }
    }

    private void BindData()
    {
        entryRepeater.DataSource = Source;
        entryRepeater.DataBind();
    }

    private void ApproveEntry(string entryId)
    {
        Entry e = plantLogService.GetEntry(entryId);
        e.IsApprove = true;
        plantLogService.UpdateEntry(e);
        Source = plantLogService.GetEntryByOwner(OwnerId);
        BindData();
    }

    private void HideEntry(string entryId)
    {
        Entry e = plantLogService.GetEntry(entryId);
        e.IsApprove = false;
        plantLogService.UpdateEntry(e);
        Source = plantLogService.GetEntryByOwner(OwnerId);
        BindData();
    }

    protected void ApproveOwnerInfo()
    {
        Owner o = plantLogService.GetOwner(OwnerId);
        o.IsApprove = true;
        plantLogService.UpdateOwner(o);
        UserInfo1.thisOwner = o;
        trApproveUserInfo.Visible = false;
        trHideUserInfo.Visible = true;
    }

    protected void HideOwnerInfo()
    {
        Owner o = plantLogService.GetOwner(OwnerId);
        o.IsApprove = false;
        plantLogService.UpdateOwner(o);
        UserInfo1.thisOwner = o;
        trApproveUserInfo.Visible = true;
        trHideUserInfo.Visible = false;
    }

    private void SetProperty()
    {
        if (WebUtility.IsAdmin(Session["memID"].ToString()))
        {
            isAdmin = true;
        }
        else
        {
            isAdmin = false;
        }
    }

    private void SetOwner()
    {
        if (OwnerId == null)
        {
            OwnerId = Request.QueryString["ownerid"];
            Owner o = plantLogService.GetOwner(OwnerId);

            if (!isAdmin)
            {
                trApproveUserInfo.Visible = false;
                trHideUserInfo.Visible = false;
            }
            else
            {
                if (o.IsApprove)
                {
                    trApproveUserInfo.Visible = false;
                }
                else
                {
                    trHideUserInfo.Visible = false;
                }
            }

            UserInfo1.thisOwner = o;
        }
    }

    private void SetListSource()
    {
        if (Source == null)
        {
            Source = plantLogService.GetEntryByOwner(OwnerId);
        }
    }

    protected void entryRepeater_ItemDataBound(object sender, RepeaterItemEventArgs e)
    {
        if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
        {
            UserControls_EntryControl entryControl = e.Item.FindControl("EntryControl1") as UserControls_EntryControl;
            entryControl.Source = e.Item.DataItem as Entry;
            entryControl.IsEditable = false;
            entryControl.IsAdmin = isAdmin;
            entryControl.IsVote = !(bool.Parse((string)ViewState["isBeforeVote"]) || bool.Parse((string)ViewState["isAfterVote"]));
            entryControl.IsOwner = false;
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
}
