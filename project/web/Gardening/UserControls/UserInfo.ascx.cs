using System;
using System.Collections;
using System.IO;
using System.Web.UI;
using Spring.Data;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Persistence;

public partial class UserControls_UserInfo : UserControl
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
            BindOwnerInfo();
    }

    private void BindOwnerInfo()
    {
        User thisUser = new User();
        thisUser =  thisTopic.Owner;
        OwnerId.Text = thisUser.UserId;
        if (thisUser.Nickname == null || thisUser.Nickname == string.Empty)
        {
            DisplayName.Text = thisUser.DisplayName[0] + "X" + thisUser.DisplayName[thisUser.DisplayName.Length - 1];
        }
        else
        {
            DisplayName.Text = thisUser.Nickname;
        }

        if (thisTopic.IsApprove)
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
        else
        {
			if(Session["FromBackEnd"] != null && Convert.ToBoolean(Session["FromBackEnd"])){
				TextTopic.Text = thisTopic.Title;
			}else{
				TextTopic.Text = "此參賽者之作品名稱已被管理員關閉，暫時不公開";
			
			}
            Description.Text = "此參賽者之作品介紹已被管理員關閉，暫時不公開";
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
