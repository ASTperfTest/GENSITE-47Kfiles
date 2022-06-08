using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using GSS.Vitals.COA.Data;

public partial class DsdASPXEdit : System.Web.UI.Page
{

    List<String> msg = new List<String>();
    private string iCUItem, iCTUnit, refID, MemberID, fileName;
    string Script = "<script>alert('連線逾時或尚未登入，請重新登入');window.close();</script>";

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["userID"] == null || string.IsNullOrEmpty(Session["userID"].ToString()))
        {
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "Noid", Script);
            return;
        }
        else
        {
            iCTUnit = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnit"].ToString();
            iCUItem = Request.QueryString["iCUItem"].ToString();
            refID = Session["ctNodeId"].ToString();
            MemberID = Session["userID"].ToString();
            if (!IsPostBack)
            {
                ControlInit();
                DataInit();
            }

        }
    }

    private void ControlInit()
    {
        txtDate.Text = DateTime.Now.Date.ToString("yyyy/M/d");

        string strQueryScript = @"SELECT deptid, deptname FROM dept ORDER BY deptid";
        DataTable dt = new DataTable();

        dt = SqlHelper.GetDataTable("ConnString", strQueryScript);
        ddlUnit.DataSource = dt;
        ddlUnit.DataTextField = "deptname";
        ddlUnit.DataValueField = "deptid";
        ddlUnit.DataBind();
    }

    private void DataInit()
    {
        string strQueryScript = @"SELECT * FROM CuDTGeneric WHERE iCUItem = @iCUItem";
        using (var reader = SqlHelper.ReturnReader("ConnString", strQueryScript,
            DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", iCUItem)))
        {
            if (reader.HasRows)
            {
                while (reader.Read())
                {
                    txtTitle.Text = reader["sTitle"].ToString();
                    txtDate.Text = Convert.ToDateTime(reader["xPostDate"].ToString()).ToString("yyyy/M/d");
                    // 是否公開DropDownList
                    if (reader["fCTUPublic"].ToString() == "Y")
                    {
                        ddlPublic.SelectedIndex = 1;
                    }
                    else if (reader["fCTUPublic"].ToString() == "N")
                    {
                        ddlPublic.SelectedIndex = 2;
                    }
                    else
                    {
                        ddlPublic.SelectedIndex = 0;
                    }
                    txtImport.Text = reader["xImportant"].ToString();
                    // 單位
                    foreach (ListItem item in ddlUnit.Items)
                    {
                        if (item.Value == reader["idept"].ToString().Trim())
                        {
                            item.Selected = true;
                            break;
                        }
                    }
                    imageNow.ImageUrl = "../public/History/" + reader["xImgFile"].ToString();
                }
            }
        }
    }

    protected void btnDelete_Click(object sender, EventArgs e)
    {
        try
        {
            string strDeleteScript = @"DELETE FROM CuDTGeneric WHERE iCUItem = @iCUItem";
            SqlHelper.ExecuteNonQuery("ConnString", strDeleteScript,
                DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", iCUItem));
            string strURL = "DsdASPXList.aspx?ItemID=" + Session["itemID"].ToString() +
                                    "&CtNodeID=" + Session["ctNodeId"].ToString();
            Response.Write("<script language='javascript'>alert('刪除完成');location.href('" + strURL + "');</script>");
        }
        catch (Exception)
        {
            throw;
        }
    }

    protected void btnConfirm_Click(object sender, EventArgs e)
    {
        check(txtTitle);
        if (msg.Count == 0)
        {
            string strUpdateScript = @"UPDATE CuDTGeneric SET sTitle = @sTitle, xPostDate = @xPostDate,
                                       fCTUPublic = @fCTUPublic, iEditor = @iEditor, iDept = @iDept,
                                       xImportant = @xImportant ";
            if (!string.IsNullOrEmpty(fileuploadImg.FileName))
            {
                string ext = System.IO.Path.GetExtension(fileuploadImg.FileName);
                Random rnd = new Random();
                // 指定路徑 ServerMapPath
                String path = Server.MapPath("../project/project/sys/public/History/");
                String fileName = "1" + DateTime.Now.ToString("MdHm") + rnd.Next(1000, 9999).ToString() + ext;
                // 儲存原始檔
                fileuploadImg.SaveAs(path + "\\" + fileName);
                strUpdateScript += ", xImgFile = @xImgFile WHERE iCUItem = @iCUItem";

                SqlHelper.ExecuteNonQuery("ConnString", strUpdateScript,
                    DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", iCUItem),
                    DbProviderFactories.CreateParameter("ConnString", "@sTitle", "@sTitle", txtTitle.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@xPostDate", "@xPostDate", txtDate.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@fCTUPublic", "@fCTUPublic", ddlPublic.SelectedValue.ToString()),
                    DbProviderFactories.CreateParameter("ConnString", "@iEditor", "@iEditor", MemberID),
                    DbProviderFactories.CreateParameter("ConnString", "@iDept", "@iDept", ddlUnit.SelectedValue.ToString()),
                    DbProviderFactories.CreateParameter("ConnString", "@xImportant", "@xImportant", txtImport.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@xImgFile", "@xImgFile", fileName));
            }
            else
            {
                strUpdateScript += " WHERE iCUItem = @iCUItem";

                SqlHelper.ExecuteNonQuery("ConnString", strUpdateScript,
                    DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", iCUItem),
                    DbProviderFactories.CreateParameter("ConnString", "@sTitle", "@sTitle", txtTitle.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@xPostDate", "@xPostDate", txtDate.Text),
                    DbProviderFactories.CreateParameter("ConnString", "@fCTUPublic", "@fCTUPublic", ddlPublic.SelectedValue.ToString()),
                    DbProviderFactories.CreateParameter("ConnString", "@iEditor", "@iEditor", MemberID),
                    DbProviderFactories.CreateParameter("ConnString", "@iDept", "@iDept", ddlUnit.SelectedValue.ToString()),
                    DbProviderFactories.CreateParameter("ConnString", "@xImportant", "@xImportant", txtImport.Text));
            }

            string strURL = "DsdASPXList.aspx?ItemID=" + Session["itemID"].ToString() +
                                    "&CtNodeID=" + Session["ctNodeId"].ToString();
            Response.Write("<script language='javascript'>alert('編修完成');location.href('" + strURL + "');</script>");
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
    protected List<String> check(TextBox sTitle)
    {
        if (sTitle.Text != String.Empty)
            return null;
        else
        {
            if (sTitle.Text == String.Empty)
                msg.Add("標題");
        }
        return msg;
    }
}