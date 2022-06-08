using System;
using System.Collections;
using System.IO;
using System.Web.UI;
using Spring.Data;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Persistence;

public partial class UserControls_TopicInfo : UserControl
{
    public Topic thisTopic
    {
        get
        {
            if (ViewState["thisTopic"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["thisTopic"] as Topic;
            }
        }
        set
        {
            ViewState["thisTopic"] = value;
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    { }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        if (WebUtility.CheckLogin(Session["memID"]))
        {
            BindOwnerInfo();
        }
    }

    private void BindOwnerInfo()
    {
        User thisUser = new User();
        thisUser = thisTopic.Owner;

        if (thisUser.UserId == Session["memID"].ToString())
        {
            TextTopic.Text = thisTopic.Title;

            if (thisTopic.Description == null)
            {
                Description.Text = "";
            }
            else
            {
                Description.Text = thisTopic.Description.Replace("\r\n", "<br />");
            }
        }

        if (thisTopic.Avatar == null || !thisTopic.IsApprove)
        {
            AvatarImage.ImageUrl = "../images/default.jpg";
        }
        else
        {
            string filePath = Server.MapPath(thisTopic.Avatar.Uri) + "shrink-" + thisTopic.Avatar.Name;
            if (File.Exists(filePath))
            {
                AvatarImage.ImageUrl = "..\\" + thisTopic.Avatar.Uri + "shrink-" + thisTopic.Avatar.Name;
            }
            else
            {
                AvatarImage.ImageUrl = "..\\" + thisTopic.Avatar.Uri + thisTopic.Avatar.Name;
            }
        }
    }
}
