using System;
using System.Data.SqlClient;

public partial class UploadAction : System.Web.UI.Page
{
    protected string _userAccount = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        _userAccount = Request.QueryString["account"];

        if (!IsPostBack)
        {
            if (_userAccount != null && _userAccount != "")
            {
                GetMemberPhoto();
            }
            else
            {
                MemberPhoto.Visible = false;
            }
        }

        if (hidFileName.Value != "" && hidFileName.Value != "undefined")
        {
            MemberPhoto.ImageUrl = hidFileName.Value;
            uploadImgBtn.Visible = false;
            deleteBtn.Visible = true;
            UploadBtn.Visible = true;
        }
        else
        {
            uploadImgBtn.Visible = true;
            UploadBtn.Visible = false;
        }
		
    }
    /// <summary>
    /// 繫結大頭貼圖片
    /// </summary>
    private void GetMemberPhoto()
    {
        string connString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;
        string fileName = string.Empty;
        string account = _userAccount;
        using (SqlConnection cn = new SqlConnection())
        {
            cn.ConnectionString = connString;
            cn.Open();

            string sql = @"SELECT photo FROM member WHERE account=@account";

            SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("account", account);

            fileName = Convert.ToString(cmd.ExecuteScalar());
        }

        string photoPath = string.Empty;

        if (fileName != "")
        {
            //讀取使用者的大頭貼
            photoPath = "/public/Profile/" + account + "/" + fileName;
            deleteBtn.Visible = true;
            
        }
        else
        {
            //讀取預設大頭貼圖片
            photoPath = "/public/Profile/" + "default.png";
            deleteBtn.Visible = false;
        }

        MemberPhoto.ImageUrl = photoPath;
    }

    protected void UploadBtn_Click(object sender, EventArgs e)
    {
		string uploadPath = Server.MapPath("Profile") + "\\" ;
		uploadPath = uploadPath.Replace("wsxd2\\Coamember","public");
		
        int limitFileSize = 1025000;//限制檔案大小為 1MB以下

        //驗證是否有檔案
        //if (this.FileUpload.HasFile)
        if (hidFileName.Value != "" && hidFileName.Value != "undefined")
        {
            //驗證檔案類型 是否為圖檔
            bool fileAllow = false;

            //string extName = System.IO.Path.GetExtension(this.FileUpload.FileName).ToLower();
            string extName = hidFileName.Value.Substring(hidFileName.Value.LastIndexOf('.')).ToLower();

            String[] allowedExtensions = { ".jpg", ".gif", ".png", ".jpeg" };
            for (int i = 0; i < allowedExtensions.Length; i++)
            {
                if (extName == allowedExtensions[i])
                {
                    fileAllow = true;
                    break;
                }
            }

            if (fileAllow)
            {
                //驗證檔案大小是否超過限制
                //if (this.FileUpload.PostedFile.ContentLength < limitFileSize)
                //{
                    try
                    {
                        string account = _userAccount;
                        uploadPath += account + "\\";

                        //判斷資料夾目錄是否存在，不存在則建立
                        if (System.IO.Directory.Exists(uploadPath) == false)
                        {
                            System.IO.Directory.CreateDirectory(uploadPath);
                        }

                        //判斷該目錄下是否存在其他檔案，存在的話就刪除
                        string[] filesOfDirectory = System.IO.Directory.GetFiles(uploadPath);

                        if (filesOfDirectory.Length > 0)
                        {
                            for (int i = 0; i < filesOfDirectory.Length; i++)
                            {
                                System.IO.File.Delete(filesOfDirectory[i]);
                            }
                        }

                        //指定完整檔案ex:mary_20090101123001.jpg
                        string fileName = account + "_" + String.Format("{0:yyyyMMddhhmmss}", System.DateTime.Now) + extName;
                        uploadPath += fileName;

                        //this.FileUpload.SaveAs(uploadPath);
                        string sourcePath = Server.MapPath(".") + @"\" + hidFileName.Value;
                        System.IO.File.Move(sourcePath, uploadPath);
                        System.IO.File.Delete(sourcePath);
                        string originalFileName = hidFileName.Value.Replace("_", "");
                        System.IO.File.Delete(Server.MapPath(".") + @"\" + originalFileName);
                        
                        //更新Member資料表，儲存 photo檔名
                        UpdateMemberTable(account, fileName);
                        //重新繫結圖片
                        GetMemberPhoto();
                    }
                    catch
                    {
                        Page.ClientScript.RegisterStartupScript(Page.GetType(), "uploadPhoto", "alert('發生錯誤，檔案無法上傳！' )", true);
                    }

                //}
                //else
                //{
                //    Page.ClientScript.RegisterStartupScript(Page.GetType(), "uploadPhoto", "alert('檔案大小不可超過1MB！' )", true);

                //}
            }
            else
            {
                Page.ClientScript.RegisterStartupScript(Page.GetType(), "uploadPhoto", "alert('上傳的檔案需為圖檔(.jpg、 .gif、 .png、 .jpeg)' )", true);
            }
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(Page.GetType(), "uploadPhoto", "alert('請選擇圖片上傳！' )", true);
        }
    }
    /// <summary>
    /// 更新Member資料表，將會員上傳的圖片檔名更新至資料表中
    /// </summary>
    /// <param name="account">會員帳號</param>
    /// <param name="photoName">圖片檔名</param>
    private void UpdateMemberTable(string account, string photoName)
    {
        string connString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

        using (SqlConnection cn = new SqlConnection())
        {
            cn.ConnectionString = connString;
            cn.Open();

            string sql = @"update member set photo=@photo where account=@account";

            SqlCommand cmd = new SqlCommand(sql, cn);
            cmd.Parameters.AddWithValue("account", account);
            if (photoName == "")
                cmd.Parameters.AddWithValue("photo", DBNull.Value);
            else
                cmd.Parameters.AddWithValue("photo", photoName);
            cmd.ExecuteNonQuery();
            
        }

        Page.ClientScript.RegisterStartupScript(Page.GetType(), "uploadPhoto", "alert('大頭照更新完成！' )", true);
    }

    protected void deleteBtn_Click(object sender, EventArgs e)
    {
        string account = _userAccount;
        if (hidFileName.Value != "" && hidFileName.Value != "undefined")
        {
            string sourcePath = Server.MapPath(".") + @"\" + hidFileName.Value;
            System.IO.File.Delete(sourcePath);
            string originalFileName = hidFileName.Value.Replace("_", "");
            System.IO.File.Delete(Server.MapPath(".") + @"\" + originalFileName);
            hidFileName.Value = "";
            uploadImgBtn.Visible = true;
            deleteBtn.Visible = false;
            UploadBtn.Visible = false;
        }

        //更新Member資料表，儲存 photo檔名
        UpdateMemberTable(account, "");
        //重新繫結圖片
        GetMemberPhoto();
    }
}