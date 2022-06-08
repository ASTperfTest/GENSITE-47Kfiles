using System;
using System.Collections;
using System.Web.UI;
using System.Web.UI.WebControls;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;

public partial class entrylist : Page
{
    private IGardeningService gardeningService;
    private bool isAdmin;
    private bool fromBackEnd = false;

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
        if (Session["memID"] != null)
        {
            SetProperty();
        }

        if (gardeningService == null)
        {
            gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        }

        SetOwner();

        //防止使用者登入後，利用自行輸入的網址想看未開放的日誌
        if (!UserInfo1.thisTopic.IsApprove)
        {
            
			if (!isAdmin)
            {
                Response.Redirect("Default.aspx");
            }
        }

        SetListSource();
        BindData();

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
        Entry e = gardeningService.GetEntry(entryId);
        e.IsApprove = true;
        gardeningService.UpdateEntry(e);
        Source = gardeningService.GetEntriesByTopic(topicId);
        BindData();
    }

    private void HideEntry(string entryId)
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
        if (topicId == null)
        {
            topicId = Request.QueryString["topicid"];
            Topic thisTopic = gardeningService.GetTopic(topicId);

            if (!isAdmin)
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

    private void SetListSource()
    {
        if (Source == null)
        {
            Source = gardeningService.GetEntriesByTopic(topicId);
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
            entryControl.IsOwner = false;
        }
    }

    private void SetViewState()
    {
        if (ViewState["hasLogin"] == null)
        {
            ViewState["hasLogin"] = WebUtility.CheckLogin(Session["memID"]).ToString();
        }
    }
}
