using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Web.UI;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;
using Gardening.Core.Persistence;
using System.Data.SqlClient;

public partial class editownerinfo : Page
{
    private IGardeningService gardeningService;
    private string[] imgExt = { ".bmp", ".gif", ".jpg", ".jpeg", ".jp2", ".mng", ".png", ".tif", ".tiff" };
    private bool fromBackEnd = false;
    private bool isNewTopic = false;
    
    private int? MaxWidth
    {
        get
        {
            return ViewState["MaxWidth"] as int?;
        }
        set
        {
            ViewState["MaxWidth"] = value;
        }
    }

    private int? MaxHeight
    {
        get
        {
            return ViewState["MaxHeight"] as int?;
        }
        set
        {
            ViewState["MaxHeight"] = value;
        }
    }

    private Topic thisTopic
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

    protected void Page_Load(object sender, EventArgs e)
    {
        SetViewState();
        string action = Request.QueryString["action"];
        if (Session["memID"] == null || Session["memID"].ToString().Trim() == string.Empty)
        {
            Response.Redirect(WebUtility.GetAppSetting("RedirectPage"));
        }

        if (gardeningService == null)
        {
            gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        }
        if (OwnerId == null)
        {
            OwnerId = Request.QueryString["ownerid"];
        }
        if (action != "" && action != null)
        {
            if (action == "new")
            {
                thisTopic = CreateTopic();
            }
        }
        else
        {
            if (topicId == null)
            {
                topicId = Request.QueryString["topicId"];
            }

            if (thisTopic == null & topicId != null)
            {
                Topic temp = gardeningService.GetTopic(topicId);
                thisTopic = temp;
            }
        }
        if (OwnerId != Session["memID"].ToString())
        {
            Response.Redirect("Default.aspx");
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
        if (Request.Form["Action"] == "Update")
        {
            SaveTopic(); 
        }

        BindData();

        if (Request.Form["Action"] == "NewImg")
        {
            ImgUp.Visible = true;
            ImgNow.Visible = false;
            NewImg.Visible = false;
            CancelNewImg.Visible = true;
        }

        if (Request.Form["Action"] == "CancelNewImg")
        {
            ImgUp.Visible = false;
            ImgNow.Visible = true;
            NewImg.Visible = true;
        }

        if (Request.Form["Action"] == "Cancel")
        {
            Cancel();
        }		
		    //管理者進來才能看到"推薦順序"
            bool isAdmin = WebUtility.IsAdmin(Session["memID"].ToString());
            if (!isAdmin)
            { 
                RecommendOrederHeader.Visible = false;
                RecommendOrederContext.Visible = false;
            }
    }

    private void BindData()
    {
        LabelOwnerId.Text = thisTopic.Owner.UserId;
        DisplayName.Text = thisTopic.Owner.DisplayName;
        NewTopic.Text = thisTopic.Title;
		RecommendOrderTextBox.Text = Convert.ToString(thisTopic.RecommendedOrder);

        if (thisTopic.Description == null)
        {
            Des.Value = "";
        }
        else
        {
            Des.Value = thisTopic.Description;
        }

        if (thisTopic.Avatar == null)
        {
            ImgUp.Visible = true;
            ImgNow.Visible = false;
            NewImg.Visible = false;
            CancelNewImg.Visible = false;
        }
        else
        {
            string filePath = Server.MapPath(thisTopic.Avatar.Uri) + "shrink-" + thisTopic.Avatar.Name;
            if (File.Exists(filePath))
            {
                ImageNow.ImageUrl = thisTopic.Avatar.Uri + "shrink-" + thisTopic.Avatar.Name;
            }
            else
            {
                ImageNow.ImageUrl = thisTopic.Avatar.Uri + thisTopic.Avatar.Name;
            }

            ImgUp.Visible = false;
            ImgNow.Visible = true;
            NewImg.Visible = true;
        }

        if (WebUtility.GetAppSetting("MaxAvatarWidth") == null)
        {
            MaxWidth = 120;
        }
        else
        {
            MaxWidth = int.Parse(WebUtility.GetAppSetting("MaxAvatarWidth"));
        }

        if (WebUtility.GetAppSetting("MaxAvatarHeight") == null)
        {
            MaxHeight = 160;
        }
        else
        {
            MaxHeight = int.Parse(WebUtility.GetAppSetting("MaxAvatarHeight"));
        }

        AvatarSize.Text = MaxWidth.Value.ToString() + " x " + MaxHeight.Value.ToString();
    }

    protected void SaveTopic()
    {
        bool success = true;
        thisTopic.Description = Server.HtmlEncode(Des.Value);
        thisTopic.Title = Server.HtmlEncode(NewTopic.Text);

        ImgFile oldFile = null;
        ImgFile iFile = null;

        // Before attempting to perform operations
        // on the file, verify that the FileUpload 
        // control contains a file.
        if (ImgUpload.HasFile)
        {
            iFile = UploadImg();

            if (iFile != null)
            {
                oldFile = thisTopic.Avatar;

                iFile.ParentId = thisTopic.TopicId;
                if (oldFile != null)
                {
                    iFile.FileId = oldFile.FileId;
                }

                thisTopic.Avatar = iFile;
            }
            else
            {
                success = false;
            }
        }

        try
        {
            thisTopic.ModifierId = thisTopic.Owner.UserId;
            thisTopic.ModifyDateTime = DateTime.Now;
            if (isNewTopic)
            {
		thisTopic.RecommendedOrder=99999;
                gardeningService.CreateTopic(thisTopic);
            }
            else
            {
                //存推薦順序
				thisTopic.RecommendedOrder = Convert.ToInt32(RecommendOrderTextBox.Text);
				gardeningService.UpdateTopic(thisTopic);
            }

            if (oldFile != null)
            {
                DeleteImgFile(oldFile);
            }
        }
        catch (Exception)
        {
            if (iFile != null)
            {
                DeleteImgFile(iFile);
            }
        }

        if (success)
        {
            Response.Redirect("uploadentry.aspx?topicid=" + thisTopic.TopicId);
        }
        else
        {
            WebUtility.WindowAlert(Page, "錯誤的檔案類型!，請選擇圖片");
        }
    }

    protected void Cancel()
    {
        if (!isNewTopic)
        {
            Response.Redirect("uploadentry.aspx?topicid=" + thisTopic.TopicId);
        }
        else
        {
            Response.Redirect("mytopiclist.aspx");
        }
    }

    private ImgFile UploadImg()
    {
        string relativePath = "Upload\\" + OwnerId + "\\avatar\\";

        // Specify the path on the server to
        // save the uploaded file to.
        String savePath = Server.MapPath(relativePath);

        if (!IsImage(Path.GetExtension(ImgUpload.FileName)) || !ImgUpload.PostedFile.ContentType.Contains("image/"))
        {
            return null;
        }

        if (!Directory.Exists(savePath))
        {
            Directory.CreateDirectory(savePath);
        }

        // Get the name of the file to upload.
        String fileName = CheckDuplicatedFileName(savePath, ImgUpload.FileName);

        ImgUpload.SaveAs(savePath + fileName);

        SaveShrinkImg(savePath, fileName);

        ImgFile img = new ImgFile();
        img.Name = fileName;
        img.Uri = relativePath;

        return img;
    }

    private void SaveShrinkImg(string tempName, string fileName)
    {
        System.Drawing.Image sourceImage = System.Drawing.Image.FromFile(tempName + fileName);

        if (sourceImage.Width > MaxWidth.Value)
        {
            System.Drawing.Image imgThumb = new System.Drawing.Bitmap(MaxWidth.Value,
                sourceImage.Height * MaxWidth.Value / sourceImage.Width);

            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(imgThumb);

            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            g.DrawImage(sourceImage, new Rectangle(0, 0, MaxWidth.Value,
                sourceImage.Height * MaxWidth.Value / sourceImage.Width), 0, 0,
                sourceImage.Width, sourceImage.Height, GraphicsUnit.Pixel);

            sourceImage.Dispose();

            imgThumb.Save(tempName + "shrink-" + fileName, ImageFormat.Jpeg);
            imgThumb.Dispose();
        }
    }

    private void DeleteImgFile(ImgFile file)
    {
        File.Delete(Server.MapPath(file.Uri) + file.Name);
    }

    private bool IsImage(string fileExt)
    {
        bool isImage = false;

        foreach (string type in imgExt)
        {

            if (type == fileExt.ToLower())
            {
                isImage = true;
                break;
            }
        }

        return isImage;
    }

    private string CheckDuplicatedFileName(string tempName, string fileName)
    {
        bool duplicted = false;

        do
        {
            duplicted = false;

            if (File.Exists(tempName + fileName))
            {
                duplicted = true;
                fileName = "_" + fileName;
            }

        } while (duplicted);

        return fileName;
    }

    private void SetViewState()
    {
        if (ViewState["hasLogin"] == null)
        {
            ViewState["hasLogin"] = WebUtility.CheckLogin(Session["memID"]).ToString();
        }
    }

    private Topic CreateTopic()
    {
        isNewTopic = true;
        Topic temp = new Topic();
        User thisUser = new User();
        thisUser.UserId = OwnerId;
        thisUser.DisplayName = Session["memName"].ToString();
        thisUser.Email = GetMail(OwnerId);
        if (Session["memNickName"] != null)
        {
            thisUser.Nickname = Session["memNickName"].ToString();
        }
        temp.Owner = thisUser;
        temp.Title = "尚未輸入作品名稱";
        temp.Description = "尚未輸入作品介紹";
        temp.CreateDateTime = DateTime.Now;
        temp.ModifierId = "System";
        temp.ModifyDateTime = DateTime.Now;
        return temp;
    }

    private void SaveNewTopic()
    {
        Topic temp = gardeningService.CreateTopic(thisTopic);
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
