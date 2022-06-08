using System;
using System.Net;
using System.IO;
using System.Data;
using System.Data.SqlClient;

public partial class download : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        DownloadFile();
    }
    private void DownloadFile()
    {
        string documentId = WebUtility.GetStringParameter("docid");
        string ver = WebUtility.GetStringParameter("ver");
        string fileName = WebUtility.GetStringParameter("name");

        AddDownloadDocTimes(documentId);

        string commonArg = "&format=" + WebUtility.GetKMAPISetting("APIFormat")
            + "&tid=0&who=" + WebUtility.GetKMAPISetting("APIActor") + "&pi=0&ps=10&api_key=" + WebUtility.GetKMAPISetting("APIKey");

        string downloadLink =
            WebUtility.GetKMAPISetting("APIUrl") + "/download/" + documentId +
            "?version_number=" + ver + "&file_name=" + fileName + commonArg;
        
        string contentType = GetFileContentType(fileName);
        string result = string.Empty;
        try
        {
            WebClient client = new WebClient();
            string browserName = Request.Browser.Browser.ToLower();
            byte[] bts = client.DownloadData(downloadLink);

            Response.Expires = 0;
            Response.ExpiresAbsolute = DateTime.UtcNow.AddYears(-1);
            Response.ContentType = contentType;
            Response.AddHeader("Content-Length", bts.Length.ToString());
            Response.AddHeader("content-disposition", "inline; filename=" + GetOfficeFileName(fileName, browserName));
            Response.BinaryWrite(bts);
        }
        catch (Exception ex)
        {
            Response.Redirect("/mp.asp?mp=1");
            Response.End();
        }
    }
    private string GetOfficeFileName(string fileName, string browserName)
    {
        FileInfo f = new FileInfo(fileName);
        string ext = f.Extension;
        browserName = browserName.ToLower();
        string result = f.Name.Replace(" ", "_");

        if (ext.Length > 0)
        {
            result = result.Replace(ext, string.Empty);
        }

        //由於檔案直接開啟時，檔名長度有所限制，為配合Excel的限制(218字元)縮短檔名
        //先扣除預設的暫存資料夾路徑長度，再扣除副檔名
        //ref:http://support.microsoft.com/kb/416351
        //ref:http://support.microsoft.com/kb/325573
        int len = 218 - ext.Length - DefaultTempFolderPath(browserName).Length;

        //當遇到長副檔名時，整體長度必定超過系統限制，因此略過裁切的動作
        if (len > 0)
        {
            switch (browserName)
            {
                case "ie":
                    result = GetLeft(result, len, true);
                    break;
                case "opera":
                    result = GetLeft(result, len, true);
                    break;
                case "applemac-safari":
                    //Chrome與Safari的名稱均為AppleMac-Safari
                    if (Request.UserAgent.ToLower().Contains("chrome"))
                    {
                        result = GetLeft(result, len, true);
                        break;
                    }
                    else
                    {
                        result = GetLeft(result, len, false);
                        break;
                    }
                case "firefox":
                    result = GetLeft(result, len, false);
                    break;
                default:
                    break;
            }
        }

        switch (browserName)
        {
            case "ie":
                return Server.UrlPathEncode(result + ext);
            case "opera":
                return Server.UrlPathEncode(result + ext);
            case "applemac-safari":
                if (Request.UserAgent.ToLower().Contains("chrome"))
                { return Server.UrlPathEncode(result + ext); }
                else
                { return result + ext; }
            case "firefox":
                return result + ext;
            default:
                return result + ext;
        }
    }
    //設定各瀏覽器預設之暫存資料夾路徑
    private string DefaultTempFolderPath(string browserName)
    {
        browserName = browserName.ToLower();
        if (browserName.Equals("ie"))
        {
            return @"C:\Documents and Settings\-xxxx-xxxx-xxxx-xxxx\Local Settings\
                    Temporary Internet Files\Content.IE5\yyyyyyyy";
        }

        if (browserName.Equals("opera"))
        {
            return @"C:\Documents and Settings\-xxxx-xxxx-xxxx-xxxx\Local Settings\
                    Application Data\Opera\Opera\profile\cache4\temporary_download";
        }

        if (browserName.Equals("applemac-safari"))
        {
            //Chrome與Safari的名稱均為AppleMac-Safari
            if (Request.UserAgent.ToLower().Contains("chrome"))
            {
                //Chrome無法直接開啟檔案，且Chrome會自動裁切檔名長度，因此空字串
                return string.Empty;
            }
            else
            {
                return @"C:\Documents and Settings\-xxxx-xxxx-xxxx-xxxx\Local Settings\Temp\yyyyyyyy.tmp";
            }
        }

        if (browserName.Equals("firefox"))
        {
            return @"C:\Documents and Settings\-xxxx-xxxx-xxxx-xxxx\Local Settings\Temp";
        }

        return string.Empty;
    }
    private string GetLeft(string source, int length, bool urlPathEncode)
    {
        if (urlPathEncode)
        {
            if (length < Server.UrlPathEncode(source).Length)
            {
                for (int i = source.Length; i > 0; i--)
                {
                    string sub = source.Substring(0, i);

                    if (Server.UrlPathEncode(sub).Length <= length)
                    {
                        return sub;
                    }
                }
            }
        }
        else
        {
            if (length < source.Length)
            {
                return source.Substring(0, length);
            }
        }

        return source;
    }

    private bool IsOfficeFile(string fileName)
    {
        FileInfo f = new FileInfo(fileName);
        switch (f.Extension.ToLower())
        {
            case ".rtf":
            case ".txt":
            case ".doc":
            case ".pot":
            case ".pts":
            case ".pps":
            case ".ppt":
            case ".csv":
            case ".xlt":
            case ".xlw":
            case ".xls":
            case ".xlsx":
            case ".xlsm":
            case ".xlsb":
            case ".xltx":
            case ".xltm":
            case ".xla":
            case ".xlam":
            case ".pptx":
            case ".pptm":
            case ".potx":
            case ".potm":
            case ".thmx":
            case ".ppsx":
            case ".ppsm":
            case ".ppam":
            case ".ppa":
            case ".docx":
            case ".docm":
            case ".dotx":
            case ".dotm":
            case ".dot":
            case ".wps":
                return true;
            default:
                return false;
        }
    }
    private string GetFileContentType(string fileName)
    {
        FileInfo fileInfo = new FileInfo(fileName);
        string contentType = string.Empty;
        switch (fileInfo.Extension.ToLower())
        {
            case ".doc":
            case ".docx":
                contentType = "application/msword";
                break;
            case ".xls":
            case ".xlt":
            case ".xlsx":
                contentType = "application/vnd.ms-excel";
                break;
            case ".ppt":
            case ".pps":
            case ".pptx":
                contentType = "application/vnd.ms-powerpoint";
                break;
            case ".pdf":
                contentType = "application/pdf";
                break;
            case ".txt":
                contentType = "text/plain";
                break;
            case ".swf":
                contentType = "application/x-shockwave-flash";
                break;
            case ".bmp":
                contentType = "image/bmp";
                break;
            case ".jpg":
            case ".jpeg":
            case ".jpe":
                contentType = "image/jpeg";
                break;
            case ".gif":
                contentType = "image/gif";
                break;
            case ".png":
                contentType = "image/png";
                break;
            case ".ico":
                contentType = "image/x-icon";
                break;
            case ".avi":
                contentType = "video/x-msvideo";
                break;
            case ".mp3":
                contentType = "audio/mpeg";
                break;
            default:
                contentType = "application/save-as";
                break;
        }
        return contentType;
    }

    /// <summary>
    /// 統計文章附件下載數
    /// </summary>
    /// <param name="ReportId"></param>
    private void AddDownloadDocTimes(string docid)
    {
        using (SqlConnection cn = new SqlConnection())
        {
            string connString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["LambdaCoaConnString"].ConnectionString;

            if (null != connString)
            {
                cn.ConnectionString = connString;
                cn.Open();

                SqlCommand cmd = new SqlCommand("SP_PerformanceStatisticsADD_DG", cn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@docId", SqlDbType.Int);
                cmd.Parameters.Add("@D", SqlDbType.Bit);
                cmd.Parameters.Add("@G", SqlDbType.Bit);
                cmd.Parameters["@docId"].Value = Convert.ToInt32(docid);
                cmd.Parameters["@D"].Value = 0;
                cmd.Parameters["@G"].Value = 1;

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw ex.GetBaseException();
                }
                finally
                {
                    cn.Close();
                }
            }
        }
    }
}
