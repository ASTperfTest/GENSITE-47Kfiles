using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;

public partial class Century_Picture_Detail : System.Web.UI.Page
{
    private int ctRootId, currentNodeId, iCTUnitPic;
    string liTemplate = @"<li><a href='/public/History/{0}'>
                        <img alt='{1}' title='{1}' src='/public/History/{0}'/>
                        <div style='text-align: center; margin: 10px; height: 30px; overflow: hidden;'>{1}</div></a></li>";

    protected void Page_Load(object sender, EventArgs e)
    {
        ctRootId = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["ctRootId"]);
        iCTUnitPic = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitPic"]);

        if (Session["memID"] != null)
        {
            string memberId = "";
            memberId = Session["memID"].ToString();
            if (Request.QueryString["kpi"] == null || Request.QueryString["kpi"] == "")
            {
                KPIBrowse kpiBrowse = new KPIBrowse(memberId, "browseInterCP", "6933");
                kpiBrowse.HandleBrowse();
                string relink = "";
                relink = "/Century/Picture_Detail.aspx?ctNodeId=" + Request.QueryString["ctNodeId"].ToString() + "&kpi=0";
                Response.Redirect(relink);
                Response.End();
            }
        }

        if (string.IsNullOrEmpty(Request.QueryString["ctNodeId"]) || Request.QueryString["ctNodeId"] == null)
        { 
            currentNodeId = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["currentNodeId"]); 
        }
        else
        {
            currentNodeId = Convert.ToInt32(Request.QueryString["ctNodeId"].ToString()); 
        }
        myDBinit();
    }

    private void myDBinit()
    {
        #region "TAGs"
        // T-SQL Recursive

        // Folder
        string strFolderScript  = "SELECT * FROM CatTreeNode WHERE  dataParent = @currentNodeId AND ctRootID = @ctRootId";

        using (var reader = SqlHelper.ReturnReader("ConnString", strFolderScript,
            DbProviderFactories.CreateParameter("ConnString", "@ctRootId", "@ctRootId", ctRootId),
            DbProviderFactories.CreateParameter("ConnString", "@currentNodeId", "@currentNodeId", currentNodeId),
            DbProviderFactories.CreateParameter("ConnString", "@ctNodeId", "@ctNodeId", currentNodeId)))
        {
            if (reader.HasRows)
            {
                labTAGs.Text = "<ul class=PicView>";
                while (reader.Read())
                {
                    labTAGs.Text += "<li><a href='Picture_Detail.aspx?ctNodeId=" + reader["ctNodeId"] + "'>" + reader["CatName"] + "</a></li>";
                }
                labTAGs.Text += "</ul>";
            }            
        }
        #endregion
                
        #region "Pictures"
        // 取10張照片
        string strPicScript = "SELECT * FROM CuDTGeneric WHERE iCTUnit = @iCTUnit and fCTUPublic = 'Y' AND refId = @currentNodeId ORDER BY newid()";

        using (var data = SqlHelper.GetDataTable("ConnString", strPicScript,
            DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnitPic),
            DbProviderFactories.CreateParameter("ConnString", "@currentNodeId", "@currentNodeId", currentNodeId)))
        {            
            // 節點下有圖片。
            if (data.Rows.Count > 0)
            {
                // temp 圖片路徑                
                litPictures.Text = "<ul class=\"album\">";
                foreach (System.Data.DataRow row in data.Rows)
                {
                    litPictures.Text += string.Format(liTemplate
                        , row["xImgFile"]
                        , row["sTitle"]); ;
                }
                litPictures.Text += "</ul>";
            }
            // 節點下無圖片；找尋該節點所包含之子節點的圖片。
            else
            {

                // 先取得ctNodeId(Request.QueryString["ctNodeId"])底下所包含之LeafNodeId
                string strctNodeIdScript = @"WITH node_Tree(ctNodeId, dataParent, CatName, DataLevel, fullPath) AS
                                 (SELECT ctNodeId, dataParent, CatName, DataLevel, cast('' as varchar(max))
                                  FROM CatTreeNode WHERE ctNodeId = @currentNodeId AND ctRootID = @ctRootId
                                        UNION ALL
                                  SELECT CatTreeNode.ctNodeId, node_Tree.ctNodeId, CatTreeNode.CatName, CatTreeNode.DataLevel
                                        ,cast(node_Tree.fullPath + '/' + CatTreeNode.CatName as varchar(max))
                                  FROM CatTreeNode INNER JOIN node_Tree ON CatTreeNode.dataParent = node_Tree.ctNodeId) ";
                string strLeafNodeIdScript = strctNodeIdScript +
                    @"
                    SELECT 
                        top 10
                        CuDTGeneric.* 
                    FROM node_Tree
                    inner join CuDTGeneric on CuDTGeneric.refId=node_Tree.ctNodeId
                    WHERE  ctNodeId <> @ctNodeId
                    order by newid()";


                using (System.Data.DataTable dt = SqlHelper.GetDataTable("ConnString", strLeafNodeIdScript,
                    DbProviderFactories.CreateParameter("ConnString", "@ctRootId", "@ctRootId", ctRootId),
                    DbProviderFactories.CreateParameter("ConnString", "@currentNodeId", "@currentNodeId", currentNodeId),
                    DbProviderFactories.CreateParameter("ConnString", "@ctNodeId", "@ctNodeId", currentNodeId)))
                {
                    litPictures.Text = "<ul class=\"album\">";

                    foreach (System.Data.DataRow GenericRow in dt.Rows)
                    {
                        litPictures.Text += string.Format(liTemplate
                            , GenericRow["xImgFile"]
                            , GenericRow["sTitle"]); ;
                    }
                    litPictures.Text += "</ul>";
                }

            }
        }
        #endregion
    }
}