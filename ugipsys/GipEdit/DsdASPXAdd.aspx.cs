using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using GSS.Vitals.COA.Data;

public partial class DsdASPXAdd : System.Web.UI.Page
{
    List<String> msg = new List<String>();
    private string iCTUnit, refID, MemberID, fileName;
    string Script = "<script>alert('連線逾時或尚未登入，請重新登入');window.close();</script>";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()) ||
            Session["ctNodeId"] == null || string.IsNullOrEmpty(Session["ctNodeId"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
        else
        {
            if (!IsPostBack)
            {
                ControlInit();
            }
            iCTUnit = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnit"].ToString();
            refID = Session["ctNodeId"].ToString();
            MemberID = Session["userID"].ToString();
        }
    }

    private void ControlInit()
    {
        txtDate.Text = DateTime.Now.Date.ToString("yyyy/M/d");
        txtImport.Text = "0";

        string strQueryScript = @"SELECT deptid, deptname FROM dept ORDER BY deptid";
        DataTable dt = new DataTable();

        dt = SqlHelper.GetDataTable("ConnString", strQueryScript);
        ddlUnit.DataSource = dt;
        ddlUnit.DataTextField = "deptname";
        ddlUnit.DataValueField = "deptid";
        ddlUnit.DataBind();
    }
    protected void btnConfirm_Click(object sender, EventArgs e)
    {

        check(fileuploadImg, txtTitle);
        if (msg.Count == 0)
        {
            try
            {
                string ext = System.IO.Path.GetExtension(fileuploadImg.FileName);
                Random rnd = new Random();
                // 指定路徑 ServerMapPath
                String path = Server.MapPath("../project/project/sys/public/History/");
                String fileName = "1" + DateTime.Now.ToString("MdHm") + rnd.Next(1000, 9999).ToString() + ext;
                // 儲存原始檔
                fileuploadImg.SaveAs(path + "\\" + fileName);


                string strInsertScript = @"INSERT INTO CuDTGeneric (iBaseDSD,iCTUnit,fCTUPublic,sTitle,iEditor,iDept
                                     , xPostDate, xImportant, refID, xImgFile) VALUES (@iBaseDSD,@iCTUnit,@fCTUPublic
                                     ,@sTitle,@iEditor, @iDept, @xPostDate, @xImportant, @refID, @xImgFile)";
                SqlHelper.ExecuteNonQuery("ConnString", strInsertScript,
                    DbProviderFactories.CreateParameter("ConnString", "@iBaseDSD", "@iBaseDSD", "7"),
                    DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnit),
                    DbProviderFactories.CreateParameter("ConnString", "@fCTUPublic", "@fCTUPublic", "Y"),
                    DbProviderFactories.CreateParameter("ConnString", "@sTitle", "@sTitle", txtTitle.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@iEditor", "@iEditor", MemberID),
                    DbProviderFactories.CreateParameter("ConnString", "@iDept", "@iDept", ddlUnit.SelectedValue.ToString()),
                    DbProviderFactories.CreateParameter("ConnString", "@xPostDate", "@xPostDate", txtDate.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@xImportant", "@xImportant", txtImport.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@refID", "@refID", refID),
                    DbProviderFactories.CreateParameter("ConnString", "@xImgFile", "@xImgFile", fileName));
                // Response.Write(fileName);
                string strURL = "DsdASPXList.aspx?ItemID=" + Session["itemID"].ToString() + 
                                    "&CtNodeID=" + Session["ctNodeId"].ToString();
                Response.Write("<script language='javascript'>alert('新增完成');location.href('"+strURL+"');</script>");
            }
            catch (Exception)
            {

                throw;
            }
        }
        else
        {
            String errMsg = String.Empty;
            for (int intX = 0; intX < msg.Count; intX++)
            {
                errMsg += msg[intX] + ",";
            }
            Response.Write("<script language='javascript'>alert('" + errMsg + "並未輸入。')</script>");
        }
    }

    // 檢查 FileUpload & TextBox
    protected List<String> check(FileUpload file, TextBox sTitle)
    {
        if (file.FileName != String.Empty && sTitle.Text != String.Empty)
            return null;
        else
        {
            if (sTitle.Text == String.Empty)
                msg.Add("標題");
            if (file.FileName == String.Empty)
                msg.Add("檔案");
        }
        return msg;
    }
}