using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Web.UI;
using PlantLog.Core;
using PlantLog.Core.Domain;
using PlantLog.Core.Service;

public partial class singleentry : Page
{
    private IPlantLogService plantLogService;
    private string[] imgExt = { ".bmp", ".gif", ".jpg", ".jpeg", ".jp2", ".mng", ".png", ".tif", ".tiff" };
    private bool fromBackEnd = false;
    
    private Entry thisEntry
    {
        get
        {
            if (ViewState["thisEntry"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["thisEntry"] as Entry;
            }
        }
        set
        {
            ViewState["thisEntry"] = value;
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

        if (thisEntry == null)
        {
            string entryId = Request.QueryString["entryid"];

            if (entryId == null)
            {
                thisEntry = new Entry();
                OwnerId = Request.QueryString["ownerid"];
            }
            else
            {
                thisEntry = plantLogService.GetEntry(entryId);
                OwnerId = thisEntry.OwnerId;
            }
        }

        if (OwnerId != Session["memID"].ToString())
        {
            Response.Redirect("Default.aspx");
        }

        SetPublicCount();

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
        if (Request.Form["Action"] == null)
        {
            DoAction("Non");
        }
        else
        {
            DoAction(Request.Form["Action"]);
        }
    }

    protected void SaveEntry()
    {
        string year = DateText.Text.Split('/')[0];
        bool success = true;

        thisEntry.Date = DateTime.Parse(DateText.Text.Replace(year, (int.Parse(year) + 1911).ToString()));
        thisEntry.Title = Server.HtmlEncode(TitleText.Text);
        thisEntry.Description = Server.HtmlEncode(Des.Value);
        thisEntry.IsPublic = bool.Parse(radioIsPublic.SelectedValue);

        if (thisEntry.EntryId == null)
        {
            thisEntry.EntryId = Utility.GetGuid();
            thisEntry.OwnerId = OwnerId;

            // Before attempting to perform operations
            // on the file, verify that the FileUpload 
            // control contains a file.
            if (ImgUpload.HasFile)
            {
                ImgFile iFile = UploadImg();

                if (iFile != null)
                {
                    thisEntry.Files = new ArrayList(new ImgFile[] { iFile });
                    thisEntry.CreatorId = thisEntry.OwnerId;
                    thisEntry.CreateDateTime = DateTime.Now;
                    thisEntry.ModifierId = thisEntry.OwnerId;
                    thisEntry.ModifyDateTime = DateTime.Now;

                    plantLogService.CreateEntry(thisEntry);
                }
                else
                {
                    success = false;
                }
            }
        }
        else
        {
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
                    oldFile = (ImgFile)thisEntry.Files[0];
                    iFile.EntryId = oldFile.EntryId;
                    iFile.FileId = oldFile.FileId;
                    thisEntry.Files = new ArrayList(new ImgFile[] { iFile });
                }
                else
                {
                    success = false;
                }
            }

            thisEntry.ModifierId = thisEntry.OwnerId;
            thisEntry.ModifyDateTime = DateTime.Now;
            plantLogService.UpdateEntry(thisEntry);

            if (oldFile != null)
            {
                DeleteImgFile(oldFile);
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
        string relativePath = "Upload\\" + OwnerId + "\\" + thisEntry.EntryId + "\\";

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

        int maxWidth;
        if (!int.TryParse(WebUtility.GetAppSetting("MaxWidth"), out maxWidth))
        {
            maxWidth = 640;
        }

        int maxHeight;
        if (!int.TryParse(WebUtility.GetAppSetting("MaxHeight"), out maxHeight))
        {
            maxHeight = 480;
        }

        if (sourceImage.Width > maxWidth)
        {
            System.Drawing.Image imgThumb = new System.Drawing.Bitmap(maxWidth,
                sourceImage.Height * maxWidth / sourceImage.Width);

            System.Drawing.Graphics g = System.Drawing.Graphics.FromImage(imgThumb);

            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

            g.DrawImage(sourceImage, new Rectangle(0, 0, maxWidth,
                sourceImage.Height * maxWidth / sourceImage.Width), 0, 0,
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

    private void SetPublicCount()
    {
        IList all = plantLogService.GetEntryByOwner(OwnerId);
        int count = 0;

        foreach (Entry e in all)
        {
            if (e.IsPublic)
            {
                count++;
            }
        }

        ClientScript.RegisterClientScriptBlock(typeof(Page), "publicCount",
            "<script language='javascript'>$(document).ready(function(){$(\"#PublicCount\").val(" +
            count.ToString() + ")});</script>");
    }

    private void DoAction(string act)
    {
        switch (act)
        {
            case "Cancel":
                Cancel();
                break;
            case "Save":
                SaveEntry();
                break;
            case "Non":

                if (thisEntry.EntryId != null)
                {
                    radioIsPublic.SelectedIndex = thisEntry.IsPublic ? 0 : 1;
                    DateText.Text = (thisEntry.Date.Year - 1911).ToString() +
                        "/" + thisEntry.Date.Month.ToString() +
                        "/" + thisEntry.Date.Day.ToString();
                    TitleText.Text = thisEntry.Title;
                    Des.Value = thisEntry.Description;

                    ImgFile img = thisEntry.Files[0] as ImgFile;
                    string filePath = Server.MapPath(img.Uri) + "shrink-" + img.Name;
                    if (File.Exists(filePath))
                    {
                        ImageNow.ToolTip = "點選觀看原圖";
                        ImageNow.ImageUrl = img.Uri + "shrink-" + img.Name;
                        ImageNow.Attributes.Add("OnClick", "javascript:window.open('" +
                            (img.Uri + img.Name).Replace("\\", "\\\\") +
                            "', '', 'menubar=no, status=no, toolbar=no, scrollbars=yes')");
                    }
                    else
                    {
                        ImageNow.ImageUrl = img.Uri + img.Name;
                    }

                    ImgUp.Visible = false;
                    ImgNow.Visible = true;
                    NewImg.Visible = true;
                }
                else
                {
                    radioIsPublic.SelectedIndex = 0;
                    DateText.Text = string.Empty;
                    TitleText.Text = string.Empty;
                    Des.Value = string.Empty;

                    ImgUp.Visible = true;
                    ImgNow.Visible = false;
                    NewImg.Visible = false;
                    CancelNewImg.Visible = false;
                }
                break;
            case "NewImg":
                ImgUp.Visible = true;
                ImgNow.Visible = false;
                NewImg.Visible = false;
                CancelNewImg.Visible = true;
                break;
            case "CancelNewImg":
                ImgUp.Visible = false;
                ImgNow.Visible = true;
                NewImg.Visible = true;
                break;
            default:
                break;
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
