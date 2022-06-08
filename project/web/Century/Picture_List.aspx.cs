using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;

public partial class Century_Picture_List : System.Web.UI.Page
{
    private int ctRootId, currentNodeId, iCTUnitPic;
    protected void Page_Load(object sender, EventArgs e)
    {
        ctRootId = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["ctRootId"]);
        iCTUnitPic = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitPic"]);
        if (string.IsNullOrEmpty(Request.QueryString["ctNodeId"]) || Request.QueryString["ctNodeId"] == null)
        { currentNodeId = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["currentNodeId"]); }
        else
        { currentNodeId = Convert.ToInt32(Request.QueryString["ctNodeId"].ToString()); }
        myDBinit();
    }

    private void myDBinit()
    {
        LitView.Text = "<ul class=\"albumFolderIcon\">";
        #region "Folder"
        // T-SQL Recursive
        string strRecursiveScript = @"WITH node_Tree(ctNodeId, dataParent, CatName, DataLevel, fullPath) AS
                                 (SELECT ctNodeId, dataParent, CatName, DataLevel, cast('' as varchar(max))
                                  FROM CatTreeNode WHERE ctNodeId = @currentNodeId AND ctRootID = @ctRootId
                                        UNION ALL
                                  SELECT CatTreeNode.ctNodeId, node_Tree.ctNodeId, CatTreeNode.CatName, CatTreeNode.DataLevel
                                        ,cast(node_Tree.fullPath + '/' + CatTreeNode.CatName as varchar(max))
                                  FROM CatTreeNode INNER JOIN node_Tree ON CatTreeNode.dataParent = node_Tree.ctNodeId) ";
        // Folder
        string strFolderScript = strRecursiveScript + "SELECT * FROM node_Tree WHERE DataLevel = 2";
        using (var reader = SqlHelper.ReturnReader("ConnString", strFolderScript,
            DbProviderFactories.CreateParameter("ConnString", "@ctRootId", "@ctRootId", ctRootId),
            DbProviderFactories.CreateParameter("ConnString", "@currentNodeId", "@currentNodeId", currentNodeId),
            DbProviderFactories.CreateParameter("ConnString", "@ctNodeId", "@ctNodeId", currentNodeId)))
        {
            string liTemplate = @"<li><a href='Picture_Detail.aspx?ctNodeId={0}'>
            <img alt='{1}' src='css/images/folder.gif' />
            <div style='text-align: center; margin: 10px; height: 30px; overflow: hidden;'>{1}</div></a></li>";
            while (reader.Read())
            {
                LitView.Text += string.Format(liTemplate
                    , reader["ctNodeId"]
                    , reader["CatName"]);
            }
        }
        LitView.EnableViewState = false;
        #endregion

        #region "Pictures"
        // 取10張照片
        /*
        string strPicScript = "SELECT TOP 10 * FROM CuDTGeneric WHERE iCTUnit = @iCTUnit AND fCTUPublic = 'Y' ORDER BY newid()";
        using (var reader = SqlHelper.ReturnReader("ConnString", strPicScript,
            DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnitPic)))
        {
            string liTemplate = @"<li><a href='/public/History/{0}'>
                <img alt='{1}' title='{1}' src='/public/History/{0}'/>
                <div style='text-align: center; margin: 10px; height: 30px; overflow: hidden;'>{1}</div></a></li>";

            while (reader.Read())
            {
                LitView.Text += string.Format(liTemplate
                    , reader["xImgFile"]
                    , reader["sTitle"]); ;
            }
        }
        LitView.Text += "</ul>";
        LitView.EnableViewState = false;
        */
        #endregion

    }
}
