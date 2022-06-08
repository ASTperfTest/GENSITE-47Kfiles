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

public partial class GIP_web_CatalogTreeUserControl : System.Web.UI.UserControl
{
	public int CurrentRootId
	{
		get { return (int)Session["User_id"]; }
		set { Session["User_id"] = value; }
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
			CatelogRepeater_DataBind();
		}
    }

	protected void CatelogRepeater_DataBind()
	{
		CatelogTreeRoot root = TopicWebHelper.getInstance().getRoot(CurrentRootId);
		IList nodes = TopicWebHelper.getInstance().getChildNodes(root);
		CatelogRepeater.DataSource = nodes;
		CatelogRepeater.DataBind();
		NodeRepeater_Refresh(CatelogRepeater);
	}

	protected void CatelogRepeater_ItemDataBound(object sender, RepeaterItemEventArgs e)
	{
		if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
		{
			CatelogTreeNode node = (CatelogTreeNode)e.Item.DataItem;

			Button MoveUpButton = (Button)e.Item.FindControl("MoveUpButton");
			MoveUpButton.CommandName = "MoveUp";
			MoveUpButton.CommandArgument = node.Id.ToString();

			Button MoveDownButton = (Button)e.Item.FindControl("MoveDownButton");
			MoveDownButton.CommandName = "MoveDown";
			MoveDownButton.CommandArgument = node.Id.ToString();

			LinkButton EditButton = (LinkButton)e.Item.FindControl("EditButton");
			EditButton.CommandName = "Edit";
			EditButton.Text = node.Name;
			if (node.Kind == CatelogTreeNode.CATELOG)
			{
				EditButton.CommandArgument = "Step3_EditFolder.aspx?nodeId=" + node.Id;
			}
			else
			{
				EditButton.CommandArgument = "Step3_EditUnit.aspx?nodeId=" + node.Id;
			}

			Repeater NodeRepeater = (Repeater)e.Item.FindControl("NodeRepeater");
			NodeRepeater.ItemDataBound += new RepeaterItemEventHandler(NodeRepeater_ItemDataBound);
			NodeRepeater.DataSource = TopicWebHelper.getInstance().getNodesByParent(node);
			NodeRepeater.DataBind();
			NodeRepeater_Refresh(NodeRepeater);
		}
	}

	protected void NodeRepeater_Refresh(Repeater repeater)
	{
		foreach (RepeaterItem item in repeater.Items)
		{
			bool isFirst = (item.ItemIndex == 0);
			bool isLast = (item.ItemIndex == repeater.Items.Count - 1);

			if (isFirst)
			{
				Button MoveUpButton = (Button)item.FindControl("MoveUpButton");
				MoveUpButton.Visible = false;
			}

			if (isLast)
			{
				Button MoveDownButton = (Button)item.FindControl("MoveDownButton");
				MoveDownButton.Visible = false;
			}
		}
	}

	protected void NodeRepeater_ItemDataBound(object sender, RepeaterItemEventArgs e)
	{
		if (e.Item.ItemType == ListItemType.Item || e.Item.ItemType == ListItemType.AlternatingItem)
		{
			CatelogTreeNode node = (CatelogTreeNode)e.Item.DataItem;

			Button MoveUpButton = (Button)e.Item.FindControl("MoveUpButton");
			MoveUpButton.CommandName = "MoveUp";
			MoveUpButton.CommandArgument = node.Id.ToString();

			Button MoveDownButton = (Button)e.Item.FindControl("MoveDownButton");
			MoveDownButton.CommandName = "MoveDown";
			MoveDownButton.CommandArgument = node.Id.ToString();

			LinkButton EditButton = (LinkButton)e.Item.FindControl("EditButton");
			EditButton.CommandName = "Edit";
			EditButton.Text = node.Name;
			if (node.Kind == CatelogTreeNode.CATELOG)
			{
				EditButton.CommandArgument = "Step3_EditFolder.aspx?nodeId=" + node.Id;
			}
			else
			{
				EditButton.CommandArgument = "Step3_EditUnit.aspx?nodeId=" + node.Id;
			}
		}
	}

	protected void CatelogRepeater_ItemCommand(object source, RepeaterCommandEventArgs e)
	{
		if (e.CommandName == "MoveUp" || e.CommandName == "MoveDown")
		{
			Repeater repeater = (Repeater)e.Item.Parent;

			int nodeId = Convert.ToInt32(e.CommandArgument);

			if (e.CommandName == "MoveUp" && e.Item.ItemIndex > 0)
			{
				TopicWebHelper.getInstance().moveUpNode(nodeId);
			}
			else if (e.CommandName == "MoveDown" && e.Item.ItemIndex < repeater.Items.Count)
			{
				TopicWebHelper.getInstance().moveDownNode(nodeId);
			}

			//CurrentCatelogId = nodeId;

			CatelogRepeater_DataBind();
		}
		else if (e.CommandName == "Edit")
		{
			Response.Redirect((string)e.CommandArgument);
		}
	}
}
