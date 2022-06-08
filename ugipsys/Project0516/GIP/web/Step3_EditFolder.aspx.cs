using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

public partial class GIP_web_Step3_EditFolder : System.Web.UI.Page
{

	public int CurrentNodeId
	{
		get { return (ViewState["CurrentNodeId"] != null ? (int)ViewState["CurrentNodeId"] : 0); }
		set { ViewState["CurrentNodeId"] = value; }
	}

    protected void Page_Load(object sender, EventArgs e)
    {
		if (!IsPostBack)
		{
			try
			{
				CurrentNodeId = Convert.ToInt32(Request.QueryString["nodeId"]);

				CatelogTreeNode node = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getCatelogFolder(CurrentNodeId);
				FolderNameTextBox.Text = node.Name;
				NodeNameMemoTextBox.Text = node.CatNameMemo;
				IsFolderOpenRadioButtonList.SelectedValue = node.InUse ? "Y" : "N";
			}
			catch (Exception ex)
			{
				ClientScript.RegisterStartupScript(Page.GetType(), "Error", "alert('目錄不存在');location.href='Step3.aspx';", true);
				Response.End();
			}
		}
    }

	protected void UpdateButton_Click(object sender, EventArgs e)
	{
		try
		{
			CatelogTreeNode node = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getCatelogFolder(CurrentNodeId);
			node.Name = FolderNameTextBox.Text;
			node.CatNameMemo = NodeNameMemoTextBox.Text;
			node.InUse = IsFolderOpenRadioButtonList.SelectedValue.Equals("Y");
			Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().updateCatelogFolder(node,Session["Name"].ToString());
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert('儲存成功');location.href='Step3.aspx';", true);
		}
		catch (Exception ex)
		{
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert(\"發生錯誤:\\n" + ex.Message + "\");", true);
		}
	}

	protected void DeleteButton_Click(object sender, EventArgs e)
	{
		try
		{
			Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().deleteCatelogFolder(CurrentNodeId);
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert('刪除成功');location.href='Step3.aspx';", true);
		}
		catch (Exception ex)
		{
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert(\"發生錯誤:\\n" + ex.Message + "\");", true);
		}
	}

	protected void AddCatelogNodeImageButton_Click(object sender, ImageClickEventArgs e)
	{
		Response.Redirect("Step3_AddUnit.aspx?parent=" + CurrentNodeId);
	}
    protected void Cancel_Click(object sender, EventArgs e)
    {
        Response.Redirect("Step3.aspx");
    }
}
