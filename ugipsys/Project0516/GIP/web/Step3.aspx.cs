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
using Microsoft.VisualBasic;

public partial class GIP_web_Step3 : System.Web.UI.Page
{
	public int CurrentRootId
	{
		get { return (int)Session["User_id"]; }
		set { Session["User_id"] = value; }
	}

    protected void Page_Load(object sender, EventArgs e)
    {
		if (!IsPostBack)
		{
			CurrentRootId = Convert.ToInt32(Session["User_id"].ToString());
		}
	}

	protected void AddCatelogFolderImageButton_Click(object sender, ImageClickEventArgs e)
	{
		Response.Redirect("Step3_AddFolder.aspx");
	}
    protected void PrevStepButton_Click(object sender, EventArgs e)
    {
        Response.Redirect("../../edit/step2_edit.aspx");
    }
    protected void NextStepButton_Click(object sender, EventArgs e)
    {
        Response.Redirect("../../homepage.aspx");
    }
    protected void CancelButton_Click(object sender, EventArgs e)
    {
        Response.Redirect("../../index.aspx");
    }
}