using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Web.UI;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;
using System.Data.SqlClient;
using System.Data;

public partial class singleentry : Page
{
    private IGardeningService gardeningService;
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

    private string TopicId
    {
        get
        {
            if (ViewState["TopicId"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["TopicId"] as string;
            }
        }
        set
        {
            ViewState["TopicId"] = value;
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        SetViewState();

        if (Session["memID"] == null || Session["memID"].ToString().Trim() == string.Empty)
        {
            Response.Redirect(WebUtility.GetAppSetting("RedirectPage"));
        }

        if (gardeningService == null)
        {
            gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        }

        OwnerId = Request.QueryString["ownerid"];
        TopicId = Request.QueryString["topicId"];

        if (thisEntry == null)
        {
            string entryId = Request.QueryString["entryid"];

            if (entryId == null)
            {
                thisEntry = new Entry();
            }
            else
            {
                thisEntry = gardeningService.GetEntry(entryId);
                OwnerId = thisEntry.CreatorId;
            }
        }

        if (OwnerId != Session["memID"].ToString() && WebUtility.IsAdmin(OwnerId))
        {
            Response.Redirect("Default.aspx");
        }
        SetPublicCount();

        if (!IsPostBack)
        {
            useManagementMasterPage.Value = fromBackEnd.ToString();
			ViewState["ReferrerUrl"] = Request.UrlReferrer.ToString();
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
			string hiddenFromBackEnd = Request.Form["ctl00$cp$useManagementMasterPage"];
            if (hiddenFromBackEnd != null && hiddenFromBackEnd != "" && Convert.ToBoolean(hiddenFromBackEnd))
            {
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
        thisEntry.TopicId = TopicId;
        if (thisEntry.EntryId == null)
        {
			//計算 kpi
			CalculateKPI();
			
            thisEntry.EntryId = Utility.GetGuid();
            thisEntry.CreatorId = OwnerId;

            // Before attempting to perform operations
            // on the file, verify that the FileUpload 
            // control contains a file.
            if (ImgUpload.HasFile)
            {
                ImgFile iFile = UploadImg();

                if (iFile != null)
                {
                    thisEntry.Files = new ArrayList(new ImgFile[] { iFile });
                    thisEntry.CreatorId = thisEntry.CreatorId;
                    thisEntry.CreateDateTime = DateTime.Now;
                    thisEntry.ModifierId = thisEntry.CreatorId;
                    thisEntry.ModifyDateTime = DateTime.Now;

                    gardeningService.CreateEntry(thisEntry);
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
                    iFile.ParentId = oldFile.ParentId;
                    iFile.FileId = oldFile.FileId;
                    thisEntry.Files = new ArrayList(new ImgFile[] { iFile });
                }
                else
                {
                    success = false;
                }
            }

            thisEntry.ModifierId = thisEntry.CreatorId;
            thisEntry.ModifyDateTime = DateTime.Now;
            gardeningService.UpdateEntry(thisEntry);

            if (oldFile != null)
            {
                DeleteImgFile(oldFile);
            }
        }

        if (success)
        {
			Response.Redirect(ViewState["ReferrerUrl"].ToString());
        }
        else
        {
            WebUtility.WindowAlert(Page, "錯誤的檔案類型!，請選擇圖片");
        }
    }
	
	//計算kpi
	private void CalculateKPI()
    { 

        using (SqlConnection tSqlConn = new SqlConnection())
        {
            string webConfigConnectionString1 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

            tSqlConn.ConnectionString = webConfigConnectionString1;
            tSqlConn.Open();

            string sql = @"SELECT memberId FROM MemberGradeShare "+  
                                            "WHERE (memberId =@memberId)  " +
                                            "AND (CONVERT(nvarchar, shareDate, 111) = CONVERT(nvarchar, GETDATE(), 111))";
            SqlCommand Cmd = new SqlCommand(sql, tSqlConn);
            Cmd.Parameters.AddWithValue("memberId", Session["memID"]);

            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommandBuilder cb = new SqlCommandBuilder();
            DataTable dt = new DataTable();

            da.SelectCommand = Cmd;
            cb.DataAdapter = da; // 指定DataAdapter
            da.Fill(dt);

            if ( dt != null && dt.Rows.Count == 1)
            {
                Cmd.CommandText= @"update MemberGradeShare " +
                                "set sharegardening = sharegardening+1" +
                                "WHERE memberId=@memberId and (CONVERT(nvarchar, shareDate, 111) = CONVERT(nvarchar, GETDATE(), 111))";
                Cmd.ExecuteNonQuery();
            }
            else if (dt != null && dt.Rows.Count == 0)
            {

                Cmd.CommandText = @"insert into MemberGradeShare " +
                                        "(memberid,sharegardening)values(@memberId,1)";
                Cmd.ExecuteNonQuery();
            }
            else
            { 
                // exception
            }
        }       
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
        IList all = gardeningService.GetEntriesByTopic(TopicId);
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
    }
}
