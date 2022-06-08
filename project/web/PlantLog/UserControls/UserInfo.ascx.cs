using System;
using System.IO;
using System.Web.UI;
using PlantLog.Core.Domain;

public partial class UserControls_UserInfo : UserControl
{
    public Owner thisOwner
    {
        get
        {
            if (ViewState["thisOwner"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["thisOwner"] as Owner;
            }
        }
        set
        {
            ViewState["thisOwner"] = value;
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
        OwnerId.Text = thisOwner.OwnerId;
        if (thisOwner.Nickname == null || thisOwner.Nickname == string.Empty)
        {
            DisplayName.Text = thisOwner.DisplayName[0] + "X" + thisOwner.DisplayName[thisOwner.DisplayName.Length - 1];
        }
        else
        {
            DisplayName.Text = thisOwner.Nickname;
        }
        TextTopic.Text = thisOwner.Topic;

        if (thisOwner.IsApprove)
        {
            TextTopic.Text = thisOwner.Topic;

            if (thisOwner.Description == null)
            {
                Description.Text = "";
            }
            else
            {
                Description.Text = thisOwner.Description.Replace("\r\n", "<br />");
            }
        }
        
        if (thisOwner.Avatar == null)
        {
            AvatarImage.ImageUrl = "../images/default.jpg";
        }
        else
        {
            string filePath = Server.MapPath(thisOwner.Avatar.Uri) + "shrink-" + thisOwner.Avatar.Name;
            if (File.Exists(filePath))
            {
                AvatarImage.ImageUrl = "..\\" + thisOwner.Avatar.Uri + "shrink-" + thisOwner.Avatar.Name;
            }
            else
            {
                AvatarImage.ImageUrl = "..\\" + thisOwner.Avatar.Uri + thisOwner.Avatar.Name;
            }
        }
    }
}
