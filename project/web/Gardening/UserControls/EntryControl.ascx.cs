using System;
using System.IO;
using System.Web.UI;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;
using System.Web;

public partial class UserControls_EntryControl : UserControl
{
    private Entry source;
    private bool isEditable;
    private bool isAdmin;
    private bool isOwner;
	protected string topicTitle;
	private IGardeningService gardeningService;
    
    public Entry Source
    {
        get
        {
            return source;
        }
        set
        {
            source = value;
        }
    }

    public bool IsEditable
    {
        get
        {
            return isEditable;
        }
        set
        {
            isEditable = value;
        }
    }

    public bool IsAdmin
    {
        get
        {
            return isAdmin;
        }
        set
        {
            isAdmin = value;
        }
    }

    public bool IsOwner
    {
        get
        {
            return isOwner;
        }
        set
        {
            isOwner = value;
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
		if (gardeningService == null)
        {
            gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        }
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        if (!isEditable)
        {
            divE.Visible = false;
        }

        if (source != null)
        {
            BindData();
			topicTitle = gardeningService.GetTopic(source.TopicId).Title;
            if (!isOwner && !isAdmin && !source.IsApprove)
            {
                data.Visible = false;
                notapprove.Visible = false;
                notpublic.Visible = false;
            }
            else
            {
                if (source.IsApprove)
                {
                    if (source.IsPublic)
                    {
                        data.Visible = true;
                        notapprove.Visible = false;
                        notpublic.Visible = false;
                    }
                    else
                    {
                        data.Visible = false;
                        notapprove.Visible = false;
                        notpublic.Visible = true;
                    }
                }
                else
                {
                    data.Visible = false;
                    notapprove.Visible = true;
                    notpublic.Visible = false;
                }
            }
        }
    }

    private void BindData()
    {
        if (isAdmin)
        {
            Hide.HRef = "javascript:HideEntry('" + source.EntryId + "');";
            Approve.HRef = "javascript:ApproveEntry('" + source.EntryId + "');";

            if (source.IsApprove)
            {
                divApprove.Visible = false;
            }
            else
            {
                divHide.Visible = false;
            }
        }
        else
        {
            divHide.Visible = false;
            divApprove.Visible = false;
        }

        Edit.HRef = "javascript:EditEntry('" + source.EntryId + "');";
        Delete.HRef = "javascript:ConfirmDelete('" + source.EntryId + "');";

        LabelDate.Text = "拍攝日期： " + (source.Date.Year - 1911).ToString() + "/"
            + source.Date.Month.ToString() + "/" + source.Date.Day.ToString();


        LabelTitle.Text = "標題： " + source.Title;
        LabelUserName.Text = "作者： " + source.CreatorId;
        string nickName = GetUserField(source.CreatorId, "nickname");
        if (nickName != null && nickName != string.Empty)
        {
            LabelUserName.Text += "(" + nickName + ")";
        }
        else
        {
            string userName = GetUserField(source.CreatorId, "realname").Trim();
            LabelUserName.Text += "(" + userName[0] + "X" + userName[userName.Length - 1] + ")";
        }
        ImgFile img = source.Files[0] as ImgFile;
        string filePath = Server.MapPath(img.Uri) + "shrink-" + img.Name;
        if (File.Exists(filePath))
        {
            ImageNow.ToolTip = "點選觀看原圖";
            ImageNow.ImageUrl = "..\\" + img.Uri + "shrink-" + img.Name;
            ImageNow.Attributes.Add("OnClick", "javascript:window.open('" +
                (img.Uri + img.Name).Replace("\\", "\\\\") +
                "', '', 'menubar=no, status=no, toolbar=no, scrollbars=yes')");
        }
        else
        {
            ImageNow.ImageUrl = "..\\" + img.Uri + img.Name;
        }

        LiteralDescription.Text = source.Description.Replace("\r\n", "<br />");
    }

    private string GetUserField(string ownerId, string fieldName)
    {
        string strSQL = "SELECT " + fieldName + " FROM Member WHERE account = '" + ownerId + "'";
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
}
