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
using Hyweb.M00.COA.GIP.TopicWeb;

public partial class GIP_web_Step3_AddFolder : System.Web.UI.Page
{
	public int CurrentRootId
	{
		get { return (ViewState["CurrentRootId"] != null ? (int)ViewState["CurrentRootId"] : 0); }
		set { ViewState["CurrentRootId"] = value; }
	}

	public int CurrentCatelogId
	{
		get { return (ViewState["CurrentCatelogId"] != null ? (int)ViewState["CurrentCatelogId"] : 0); }
		set { ViewState["CurrentCatelogId"] = value; }
	}

    protected void Page_Load(object sender, EventArgs e)
    {
		if (!IsPostBack)
		{
			CurrentRootId = Convert.ToInt32(Request.QueryString["rootId"]);
			CurrentCatelogId = Convert.ToInt32(Request.QueryString["catelogId"]);
		}
    }

	protected void InsertButton_Click(object sender, EventArgs e)
	{
        int rootId = Convert.ToInt32(Session["User_id"].ToString());
		int parentId = 0;

		if (HasChildRadioButtonList.Text.Equals("Y"))
		{
			CatelogTreeNode node = TopicWebHelper.getInstance().addCatelogFolder(FolderNameTextBox.Text, IsFolderOpenRadioButtonList.SelectedValue.Equals("Y"), rootId, parentId, Session["Name"].ToString(),NodeNameMemoTextBox.Text);
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert(\"新增完成\");location.href=\"Step3.aspx\";", true);
		}
		else
		{
			Response.Redirect("Step3_AddUnit.aspx?name=" + FolderNameTextBox.Text + "&open=" + IsFolderOpenRadioButtonList.SelectedValue+"&level=1&memo=" + NodeNameMemoTextBox.Text);
		}
	}

    protected void cancelButton_Click(object sender, EventArgs e)
    {
        Response.Redirect("Step3.aspx");
    }
}
