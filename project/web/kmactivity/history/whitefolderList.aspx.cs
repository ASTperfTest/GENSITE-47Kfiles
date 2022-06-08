using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Collections;
using GSS.Vitals.COA.Data;
using System.Data;

public partial class kmactivity_history_whitefolderList : System.Web.UI.Page
{
    DataTable data;
    int iCTUnitPic = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitPic"]);
    protected string errmsg = string.Empty;

    string selectNodestring = @"
            with notTree(sourceId ,parentId ,ctNodeId, catName, fullpath) as
            (
	            select 
		             ctNodeId
		            ,DataParent
		            ,ctNodeId
		            ,catName
		            ,cast(catName as nvarchar(max))
	            from mGIPcoanew..CatTreeNode where CtNodeID in (
	            select * from HISTORY_PICTURE..Picture_node)
	            union all
	            select 		 
		             sourceId
		            ,CatTreeNode.DataParent
		            ,CatTreeNode.ctNodeId
		            ,CatTreeNode.catName
		            ,CatTreeNode.catName + '/' + notTree.fullpath
	            from mGIPcoanew..CatTreeNode as CatTreeNode
	            inner join notTree on notTree.parentId=CatTreeNode.ctNodeId
            )
            select sourceId as NodeId, Fullpath from notTree where parentId = 0;";


    private void BindData()
    {

        data = SqlHelper.GetDataTable("ODBCDSN", selectNodestring);
        GridView1.DataSource = data;
        GridView1.DataBind();
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        BindData();
    }
    protected void GridView1_RowDeleting(object sender, GridViewDeleteEventArgs e)
    {
        int nodeId = int.Parse(GridView1.Rows[e.RowIndex].Cells[0].Text);

        string str = "delete from HISTORY_PICTURE..picture_node where nodeId = @CtNodeID";

        SqlHelper.ExecuteNonQuery("ODBCDSN", str,
                DbProviderFactories.CreateParameter("ConnString", "@CtNodeID", "@CtNodeID", nodeId));

        errmsg = "節點[" + nodeId.ToString() + "] 刪除成功!!";

        BindData();
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        int nodeId = 0;
        if (int.TryParse(this.TextBox1.Text, out nodeId) && !string.IsNullOrEmpty(this.TextBox1.Text))
        {
            string sql = @"
            set nocount on 
            if not exists(select * from mGIPcoanew..CatTreeNode where CtNodeID = @CtNodeID and ctUnitId=@ctUnitId)
            begin
	            select 'N'
            end
            else
            begin
	            insert into HISTORY_PICTURE..picture_node (nodeId)
	            select @CtNodeID
	            where not exists (select * from HISTORY_PICTURE..picture_node where nodeId = @CtNodeID)

	            select 'Y'
            end";

            string checker = SqlHelper.ReturnScalar("ODBCDSN", sql,
                DbProviderFactories.CreateParameter("ConnString", "@CtNodeID", "@CtNodeID", nodeId),
                DbProviderFactories.CreateParameter("ConnString", "@ctUnitId", "@ctUnitId", iCTUnitPic)).ToString();

            if (checker == "Y")
            {
                errmsg = "節點[" + nodeId.ToString() + "]加入成功!!";
            }
            else
            {
                errmsg = "節點代碼錯誤!!" ;
            }
            BindData();
            return;
        }

        errmsg = "節點代碼錯誤!";
    }
}
