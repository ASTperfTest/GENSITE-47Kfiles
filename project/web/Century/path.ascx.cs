using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;

public partial class path : System.Web.UI.UserControl
{
    int ctNodeId;
    protected void Page_Load(object sender, EventArgs e)
    {

        string CatName = string.Empty;
        if (Request.QueryString["ctNodeId"] != null &&
                    (!string.IsNullOrEmpty(Request.QueryString["ctNodeId"].ToString())))
        {
            ctNodeId = int.Parse(Request.QueryString["ctNodeId"]);
            CatName = GetPathByCtNodeId(ctNodeId);
        }
        string month = string.Empty;
        if (Request.QueryString["month"] != null &&
                    (!string.IsNullOrEmpty(Request.QueryString["month"].ToString())))
        {
            month = GetPathByMonth(Request.QueryString["month"].ToString());
        }

        // 透過檔案取得檔名
        string strPhysicalPath = System.IO.Path.GetFileName(Request.PhysicalPath);
        labPath.Text = "<ul id='path_menu'>";
        labPath.Text += "<li><a href='../mp.asp?mp=1'>首頁</a></li>";
        labPath.Text += "<li style='top: 10px;'>></li>";
        labPath.Text += "<li><a href='#'>百年農業發展史</a></li>";
        switch (strPhysicalPath)
        {
            case "Event_List.aspx":
                labPath.Text += "<li style='top: 10px;'>></li>";
                labPath.Text += "<li><a href='Event_List.aspx'>大事紀</a></li>";
                if (!string.IsNullOrEmpty(month))
                {
                    labPath.Text += "<li style='top: 10px;'>></li>";
                    labPath.Text += "<li><a href='Event_List.aspx?month="
                        + Request.QueryString["month"].ToString() + "'>" + month + "</a></li>";
                }
                break;
            case "History_List.aspx":
                labPath.Text += "<li style='top: 10px;'>></li>";
                labPath.Text += "<li><a href='History_List.aspx'>歷史上的今天</a></li>";
                break;
            case "Picture_List.aspx":
                labPath.Text += "<li style='top: 10px;'>></li>";
                labPath.Text += "<li><a href='Picture_List.aspx'>珍貴老照片</a></li>";
                break;
            case "Picture_Detail.aspx":
                PrintCurrentPicPath(ctNodeId
                    , Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["ctRootId"])
                    , ref labPath);
                break;
            case "Story_List.aspx":
            case "Story_Detail.aspx":
                labPath.Text += "<li style='top: 10px;'>></li>";
                labPath.Text += "<li><a href='Story_List.aspx'>農業故事</a></li>";
                break;
            default:
                break;
        }
        labPath.Text += "</ul>";
    }

    // 傳回CatName
    protected string GetPathByCtNodeId(int ctNodeId)
    {
        string strQueryScript = @"SELECT CatName FROM CatTreeNode WHERE CtNodeId = @CtNodeId";
        using (var reader = SqlHelper.ReturnReader("ConnString", strQueryScript,
            DbProviderFactories.CreateParameter("ConnString", "@CtNodeId", "@CtNodeId", ctNodeId)))
        {
            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    return reader["CatName"].ToString();
                }
            }
        }
        return null;
    }

    // 傳回月份
    protected string GetPathByMonth(string month)
    {
        string strPath = string.Empty;
        switch (month)
        {
            case "1":
                strPath = "1月";
                return strPath;
            case "2":
                strPath = "2月";
                return strPath;
            case "3":
                strPath = "3月";
                return strPath;
            case "4":
                strPath = "4月";
                return strPath;
            case "5":
                strPath = "5月";
                return strPath;
            case "6":
                strPath = "6月";
                return strPath;
            case "7":
                strPath = "7月";
                return strPath;
            case "8":
                strPath = "8月";
                return strPath;
            case "9":
                strPath = "9月";
                return strPath;
            case "10":
                strPath = "10月";
                return strPath;
            case "11":
                strPath = "11月";
                return strPath;
            case "12":
                strPath = "12月";
                return strPath;
            default:
                return strPath;
        }
    }



    private void PrintCurrentPicPath(int ctNodeId, int ctRootId, ref Label labPath)
    {
        string strRecursiveScript = @"
            WITH node_Tree(ctNodeId, dataParent, CatName, DataLevel, fullPath) AS
            (
            SELECT ctNodeId
	            , dataParent
	            , CatName
	            , DataLevel
	            , cast('' as varchar(max))
            FROM CatTreeNode WHERE ctNodeId = @ctNodeId
            UNION ALL
            SELECT 
	              CatTreeNode.ctNodeId
	            , CatTreeNode.dataParent
	            , CatTreeNode.CatName
	            , CatTreeNode.DataLevel
                ,cast(CatTreeNode.CatName + '/' + node_Tree.fullPath as varchar(max))
            FROM CatTreeNode 
            INNER JOIN node_Tree ON CatTreeNode.ctNodeId = node_Tree.dataParent
            where CatTreeNode.ctNodeId != @ctRootId
            )
            select * from node_Tree
            order by datalevel";


        using (var reader = SqlHelper.ReturnReader("ConnString", strRecursiveScript,
            DbProviderFactories.CreateParameter("ConnString", "@ctRootId", "@ctRootId", ctRootId),
            DbProviderFactories.CreateParameter("ConnString", "@ctNodeId", "@ctNodeId", ctNodeId)))
        {
            string liTemplate = @"<li style='top: 10px;'>></li>
            <li><a href='Picture_Detail.aspx?ctNodeId={0}'>{1}</a></li>";

            string liTemplateLV1 = @"<li style='top: 10px;'>></li>
            <li><a href='Picture_List.aspx?&mp=1&{0}'>{1}</a></li>";

            while (reader.Read())
            {
                if (reader["DataLevel"].ToString() == "1")
                {
                    labPath.Text += string.Format(liTemplateLV1
                        , reader["ctNodeId"]
                        , reader["CatName"]);
                }
                else
                {
                    labPath.Text += string.Format(liTemplate
                        , reader["ctNodeId"]
                        , reader["CatName"]);
                }
            }
        }
    }

}