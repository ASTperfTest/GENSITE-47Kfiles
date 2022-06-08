using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Web.UI;
using PlantLog.Core;
using PlantLog.Core.Domain;
using PlantLog.Core.Service;

public partial class editownerinfo : Page
{
    private IPlantLogService plantLogService;
    private string[] imgExt = { ".bmp", ".gif", ".jpg", ".jpeg", ".jp2", ".mng", ".png", ".tif", ".tiff" };
    private bool fromBackEnd = false;
    
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

    private Owner thisOwner
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

    protected void Page_Load(object sender, EventArgs e)
    {
        SetViewState();

        if (Session["memID"] == null || Session["memID"].ToString().Trim() == string.Empty)
        {
            Response.Redirect(WebUtility.GetAppSetting("RedirectPage"));
        }

        if (plantLogService == null)
        {
            plantLogService = Utility.ApplicationContext["PlantLogService"] as IPlantLogService;
        }

        if (thisOwner == null)
        {
            OwnerId = Request.QueryString["ownerid"];
            thisOwner = plantLogService.GetOwner(OwnerId);
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

        if (fromBackEnd)
        {
            this.MasterPageFile = WebUtility.managementMasterPageFile;
        }
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        if (Request.Form["Action"] == "Update")
        {
            SaveOwner();
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
    }

    private void BindData()
    {
        LabelOwnerId.Text = thisOwner.OwnerId;
        DisplayName.Text = thisOwner.DisplayName;
        NewTopic.Text = thisOwner.Topic;

        if (thisOwner.Description == null)
        {
            Des.Value = "";
        }
        else
        {
            Des.Value = thisOwner.Description;
        }

        if (thisOwner.Avatar == null)
        {
            ImgUp.Visible = true;
            ImgNow.Visible = false;
            NewImg.Visible = false;
            CancelNewImg.Visible = false;
        }
        else
        {
            string filePath = Server.MapPath(thisOwner.Avatar.Uri) + "shrink-" + thisOwner.Avatar.Name;
            if (File.Exists(filePath))
            {
                ImageNow.ImageUrl = thisOwner.Avatar.Uri + "shrink-" + thisOwner.Avatar.Name;
            }
            else
            {
                ImageNow.ImageUrl = thisOwner.Avatar.Uri + thisOwner.Avatar.Name;
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

    protected void SaveOwner()
    {
        bool success = true;

        thisOwner.Description = Server.HtmlEncode(Des.Value);
        thisOwner.Topic = Server.HtmlEncode(NewTopic.Text);

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
                oldFile = thisOwner.Avatar;

                iFile.EntryId = OwnerId;
                if (oldFile != null)
                {
                    iFile.FileId = oldFile.FileId;
                }

                thisOwner.Avatar = iFile;
            }
            else
            {
                success = false;
            }
        }

        try
        {
            thisOwner.ModifierId = thisOwner.OwnerId;
            thisOwner.ModifyDateTime = DateTime.Now;
            plantLogService.UpdateOwner(thisOwner);

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
            Response.Redirect("uploadentry.aspx");
        }
        else
        {
            WebUtility.WindowAlert(Page, "錯誤的檔案類型!，請選擇圖片");
        }
    }

    protected void Cancel()
    {
        Response.Redirect("uploadentry.aspx");
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
