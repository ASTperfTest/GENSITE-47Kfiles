using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing.Imaging;
using System.Drawing;
using System.Data.SQLite;
using System.Text;
using System.Data;

public partial class editlogo : System.Web.UI.Page
{
    protected static string SubjectId = "";
    private static string creatorName = "";
    private string currentFolderPath = System.Configuration.ConfigurationManager.AppSettings.GetValues("Path")[0];
    private string DBPath = System.Configuration.ConfigurationSettings.AppSettings["DB"].ToString();
    private string type = System.Configuration.ConfigurationSettings.AppSettings["Type"].ToString();
    protected string title = System.Configuration.ConfigurationSettings.AppSettings["TypeName"].ToString();
    
    protected void Page_Load(object sender, EventArgs e)
    {
        string nowdate = System.DateTime.Now.ToShortDateString();
        DateTime datenow = Convert.ToDateTime(nowdate);
        DateTime deadline = Convert.ToDateTime(System.Configuration.ConfigurationSettings.AppSettings["Deadline"].ToString());
        if (DateTime.Compare(datenow, deadline) > 0)
        {
            Response.Write("<script language=javascript>alert('活動已結束');location.href='http://kminter.coa.gov.tw';window.close();</script>");
            return ;
        }
	DateTime startDay=Convert.ToDateTime(System.Configuration.ConfigurationSettings.AppSettings["startDay"].ToString());
        if (DateTime.Compare(startDay,datenow) > 0)
        {
            Response.Write("<script language=javascript>alert('活動尚未開始');location.href='http://kminter.coa.gov.tw';window.close();</script>");
            return ;
        }
	if (Session["memID"] != null && Session["memID"] != ""){
		SubjectId= Session["memID"].ToString();
		creatorName=Session["memName"].ToString();
		GetMatchLogo();
	}else{
	    Response.Write("<script language=javascript>alert('請先登入會員');location.href='http://kminter.coa.gov.tw';window.close();</script>");
            return ;
	}
    }

    //檢查版本A是否存在 有圖才顯示相關資訊並且修改botton text
    //當A有圖在秀出B的按鈕
    private void GetMatchLogo()
    {
        matchinfo.Visible = true;
        DataTable db = new DataTable();
        db = SelectData(SubjectId);
        if (db.Rows.Count >= 1)
        {
	    if (System.IO.File.Exists(db.Rows[0]["IMAGE_PATH"].ToString()))
            {
				string extension=System.IO.Path.GetFileName(db.Rows[0]["IMAGE_PATH"].ToString());
                ImageA.ImageUrl = "/logoSelection/download.aspx?fileName="+extension;
                ImageA.Visible = true;
                UploadlogoA.Text = "重新上傳";
            }
            else
            {
                ImageA.Visible = false;
                UploadlogoA.Text = "上傳";
            }
        }
        else
        {
            ImageA.Visible = false;
            UploadlogoA.Text = "上傳";
        }
    }

    protected void UploadLogoA_Click(object sender, EventArgs e)
    {
        string result = string.Empty;
        if (FileUploadA.HasFile)
        {
            string fileExtension = System.IO.Path.GetExtension(FileUploadA.FileName);
            if (CheckFileExtension(fileExtension))
            {
                result = SubjectId + fileExtension;
                GetUpdate(result);
                FileUploadA.PostedFile.SaveAs(currentFolderPath + result); 
                GetMatchLogo();
            }
            else
            {
                Response.Write("<script language=javascript>alert('檔案格式不正確,必須是jpg')</script>");
                return;
            }
        }
        else
        {
            Response.Write("<script language=javascript>alert('請輸入檔案')</script>");
        }
    }

    private void GetUpdate(string result)
    {
        DataTable db = SelectData(SubjectId);
        DateTime datetemp = DateTime.Now;
        if (db.Rows.Count > 0)
        {
            System.IO.File.Delete(db.Rows[0]["IMAGE_PATH"].ToString());
            UpdateDatatable("updata", SubjectId, currentFolderPath + result, datetemp.ToString());
        }
        else
        {
            UpdateDatatable("insert", SubjectId, currentFolderPath + result, datetemp.ToString());
        }
    }

    private static DataTable SelectData(string id)
    {
        DataTable dt = new DataTable();
        try
        {
            string dbpatch = System.Configuration.ConfigurationSettings.AppSettings["DB"].ToString();
            int type =int.Parse(System.Configuration.ConfigurationSettings.AppSettings["Type"].ToString());
            SQLiteConnection cnn = new SQLiteConnection("Data Source=" + dbpatch);
            cnn.Open();
            SQLiteCommand mycommand = new SQLiteCommand(cnn);
            SQLiteParameter CREATOR = new SQLiteParameter("@CREATOR");
            SQLiteParameter activeType = new SQLiteParameter("@TYPE");
            mycommand.Parameters.Add(CREATOR);
            mycommand.Parameters["@CREATOR"].Value =id;
            mycommand.Parameters.Add(activeType);
            mycommand.Parameters["@TYPE"].Value = type;
            mycommand.CommandText = "select UID,CREATOR,IMAGE_PATH from LOGO Where CREATOR = @CREATOR AND TYPE = @TYPE";
            SQLiteDataReader reader = mycommand.ExecuteReader();
            dt.Load(reader);
            reader.Close();
            cnn.Close();
        }
        catch
        {

        }
        return dt;
    }

    private  DataTable UpdateDatatable(string cmd,string id,string patch,string dateString)
    {
        string dbpatch = System.Configuration.ConfigurationSettings.AppSettings["DB"].ToString();
        int type = int.Parse(System.Configuration.ConfigurationSettings.AppSettings["Type"].ToString());
        DataTable dt = new DataTable();

        try
        {
            SQLiteConnection cnn = new SQLiteConnection("Data Source=" + dbpatch);
            cnn.Open();

            SQLiteCommand mycommand = new SQLiteCommand(cnn);
            SQLiteParameter CREATOR = new SQLiteParameter("@CREATOR");
            SQLiteParameter activeType = new SQLiteParameter("@TYPE");
            SQLiteParameter date = new SQLiteParameter("@DATETIME");
            SQLiteParameter imagePath = new SQLiteParameter("@IMAGEPATH");
            SQLiteParameter dcoComplete = new SQLiteParameter("@COMPLETED");
	    SQLiteParameter creatorDisplayName= new SQLiteParameter("@CREATOR_DISPLAY_NAME");
            	
            mycommand.Parameters.Add(CREATOR);
            mycommand.Parameters["@CREATOR"].Value = id;
            mycommand.Parameters.Add(activeType);
            mycommand.Parameters["@TYPE"].Value = type;
            mycommand.Parameters.Add(date);
            mycommand.Parameters["@DATETIME"].Value = dateString;
            mycommand.Parameters.Add(imagePath);
            mycommand.Parameters["@IMAGEPATH"].Value = patch;
            mycommand.Parameters.Add(dcoComplete);
            mycommand.Parameters["@COMPLETED"].Value = false;
	    mycommand.Parameters.Add(creatorDisplayName);
            mycommand.Parameters["@CREATOR_DISPLAY_NAME"].Value = creatorName;
            
	    if (cmd == "updata")
            {
                mycommand.CommandText = "update LOGO Set IMAGE_PATH =@IMAGEPATH,LAST_MODIFIER =@CREATOR,LAST_MODIFY_DATETIME = @DATETIME where  CREATOR = @CREATOR AND TYPE=@TYPE";
            }
            else if (cmd == "insert")
            {
			    mycommand.CommandText = "insert into LOGO (CREATION_DATETIME,CREATOR,CREATOR_DISPLAY_NAME,IMAGE_PATH,LAST_MODIFIER,LAST_MODIFY_DATETIME,TYPE,COMPLETED) values(@DATETIME,@CREATOR,@CREATOR_DISPLAY_NAME,@IMAGEPATH,@CREATOR,@DATETIME,@TYPE,@COMPLETED)";
            }
            SQLiteDataReader reader = mycommand.ExecuteReader();
            dt.Load(reader);
            reader.Close();
            cnn.Close();
        }
        catch (Exception ex)
        {
			throw ex;
        }
        return dt;
    }

    private Boolean CheckFileExtension(string fileExtension)
    {
        string[] extensions = new string[] { ".jpg",".jpeg"};
        foreach (string str in extensions)
        {
            if (string.Compare(fileExtension,str.ToLower(),true) == 0)
            {
                return true;
            }
        }

        return false;
    }

}
