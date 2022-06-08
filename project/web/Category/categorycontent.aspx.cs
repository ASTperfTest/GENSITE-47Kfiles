#region Imports
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using GSS.Vitals.API.Client.Contracts.KnowledgeBase;
using System.Collections;
using System.Text;
using GSS.Vitals.API.Client;
using GSS.Vitals.API.Client.Contracts.Index;
using System.Collections.Specialized;
using GSS.Vitals.API;
using System.Net;
using Jayrock.Json;
using Jayrock.Json.Conversion;
using System.Data;
using System.Data.SqlClient;
#endregion

public partial class Category_categorycontent : System.Web.UI.Page
{
    #region Fields
    private RestRpcClient api = null;
    private string UrlStr = string.Empty;
    protected string relationsArticle = string.Empty;
    #endregion

    #region Page Event
    protected void Page_Load(object sender, EventArgs e)
    {
        string clientIp = Request.ServerVariables["REMOTE_ADDR"];
        if (!string.IsNullOrEmpty(clientIp) && clientIp == "172.16.18.56")
        {
            Response.Redirect("http://www.coa.gov.tw");
            Response.End();
        }
        string CategoryId = WebUtility.GetStringParameter("CategoryId");
        string ActorType = WebUtility.GetStringParameter("ActorType");
        string ReportId = WebUtility.GetStringParameter("ReportId"); 
		api = WebUtility.GetAPIClient();

        if (string.IsNullOrEmpty(ActorType)) ActorType = "002";//如果沒帶值，就用'消費者知識樹'
        if (ActorType == "004") ActorType = "003";

        DocumentDetailInfo documentDetailInfo = null;
        try
        {
            documentDetailInfo = api.documents.Get(int.Parse(ReportId));
        }
        catch (Exception ex)
        {
            Response.Redirect("/Category/categorylist.aspx?CategoryId=&ActorType=000&t=a");
            Response.End();
        }

        CategoryInfo[] cai = documentDetailInfo.Categories;
		//檢查是否為農業知識樹下的分類
        //暫時處置方式以後須修改
        bool validCategory = false;
        validCategory = IsKmwebCategory(cai);

        if (validCategory == false)
        {
            Response.Redirect("/Category/categorylist.aspx?CategoryId=&ActorType=000&t=a");
            Response.End();
        }


        //傳入的分類可能不是在農業知識樹下
        validCategory = false;
		if (!string.IsNullOrEmpty(CategoryId))
		{
			CategoryInfo[] paths = api.categories.GetPath(int.Parse(CategoryId));
			if (paths.Length > 0)
			{
				foreach (CategoryInfo categoryInfo in paths)
				{
					if (categoryInfo.CategoryId.ToString() == WebUtility.GetAppSetting("CategoryTreeRootNode"))
					{
						validCategory = true;
						break;
					}
				}
			}
		}

        if (!validCategory)
		{
			//從目前文件取得第一個農業知識樹分類
			bool continueFlag = true;
			if (documentDetailInfo.Categories != null && documentDetailInfo.Categories.Length > 0)
			{
				for (int i = 0; i < documentDetailInfo.Categories.Length; i++)
				{
					if(continueFlag)
					{
						CategoryInfoCollection categoryInfos = documentDetailInfo.Categories[i].Path;
						foreach (CategoryInfo categoryInfo in categoryInfos)
						{
							if (continueFlag && (categoryInfo.CategoryId == int.Parse(WebUtility.GetAppSetting("CategoryTreeRootNode"))))
							{ 
								CategoryId = categoryInfos[categoryInfos.Count-1].CategoryId.ToString();
								continueFlag = false;
							}
						}
					}
				}
			}			
		}
			
        if (WebUtility.GetStringParameter("kpi") != "0")
        {
            string relink = "/category/categorycontent.aspx?ReportId="
                + ReportId
                + "&CategoryId="
                + CategoryId
                + "&ActorType="
                + ActorType + "&kpi=0";
            Response.Redirect(relink);
            Response.End();
        }
        
        #region 設定Tab區塊
        string link = "";
        link = "categorylist.aspx?CategoryId={catid}&ActorType={actortype}&t=a";
        TabText.Text = "<ul class=group>";
        switch (ActorType)
        {
            case "000":
                NavUrlText.Text = "<a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", ActorType) + "\" title=\"知識庫首頁\">知識庫首頁</a>";
                NavTitleText.Text = "知識庫首頁";

                TabText.Text += "<li class=active><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "000") + "\">知識庫首頁</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "002") + "\">消費者</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "001") + "\">生產者</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "003") + "\">學者</a></li>";

                break;
            case "001":
                NavUrlText.Text = "<a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", ActorType) + "\" title=\"生產者知識庫\">生產者知識庫</a>";
                NavTitleText.Text = "生產者知識庫";

                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "000") + "\">知識庫首頁</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "002") + "\">消費者</a></li>";
                TabText.Text += "<li class=active><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "001") + "\">生產者</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "003") + "\">學者</a></li>";

                break;
            case "002":
                NavUrlText.Text = "<a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", ActorType) + "\" title=\"消費者知識庫\">消費者知識庫</a>";
                NavTitleText.Text = "消費者知識庫";

                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "000") + "\">知識庫首頁</a></li>";
                TabText.Text += "<li class=active><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "002") + "\">消費者</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "001") + "\">生產者</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "003") + "\">學者</a></li>";

                break;
            case "003":
                NavUrlText.Text = "<a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", ActorType) + "\" title=\"學者知識庫\">學者知識庫</a>";
                NavTitleText.Text = "學者知識庫";

                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "000") + "\">知識庫首頁</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "002") + "\">消費者</a></li>";
                TabText.Text += "<li><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "001") + "\">生產者</a></li>";
                TabText.Text += "<li class=active><a href=\"" + link.Replace("{catid}", "").Replace("{actortype}", "003") + "\">學者</a></li>";

                break;
            default:
                Response.Redirect("/mp.asp?mp=1");
                Response.End();
                break;
        }

        TabText.Text += "</ul>";
        #endregion

        #region 設定左邊TreeView
        UrlStr = "categorylist.aspx?CategoryId={0}&ActorType={1}";

        XmlDocument doc = new XmlDocument();
        ArrayList data = new ArrayList();
        WebUtility.GenerateXMLDataSource(doc, int.Parse(CategoryId), NavTitleText.Text, UrlStr, data);

        CatTreeViewDS.Data = doc.InnerXml;
        CatTreeView.DataSourceID = "CatTreeViewDS";
        #endregion

        #region 設定SiteMap
        //#取得目前節點分類資訊
        CategoryInfo[] currentCategoryInfos = api.categories.GetPath(int.Parse(CategoryId));
        //#取得目前節點的路徑,在依照路作出SiteMap
        if (currentCategoryInfos.Length > 0)
        {
            foreach (CategoryInfo categoryInfo in currentCategoryInfos)
            {
                if (categoryInfo.CategoryId.ToString() != WebUtility.GetAppSetting("CategoryTreeRootNode")
                   && categoryInfo.DisplayName != "#root#")
                {
                    NavUrlText.Text += @"> "
                        + @"<a href= "
                        + link.Replace("{catid}",
                        categoryInfo.CategoryId.ToString()).Replace("{actortype}", ActorType)
                        + @" title= " + categoryInfo.DisplayName + @">"
                        + categoryInfo.DisplayName + @"</a>";
                }
            }
        }
        #endregion

        #region 設定文章內容
        StringBuilder sb = new StringBuilder();
        link = "";
        //#取出Document
        ExtendedSearchResultPagedCollection result = api.search.ByAdvanced("_l_unique_key:" + ReportId,
                    null, null, null,
                    null, string.Empty, null, null, string.Empty, null, false, false, false, false, SpecialFields.FieldType.Clix);

        SearchResult sr = new SearchResult();

        SourceFileInfo[] attachmentFiles = null;
        if (documentDetailInfo != null)
        {
            if (Session["BrowseTitleNew"] != documentDetailInfo.VersionTitle)
            {
                Session["BrowseTitleNew"] = documentDetailInfo.VersionTitle;
                Session["BrowseLinkNew"] = "/category/categorycontent.aspx?ReportId="
                    + ReportId
                    + "&CategoryId="
                    + CategoryId
                    + "&ActorType="
                    + ActorType;
            }

            sb.AppendLine("<ul class=\"Function2\" xmlns:hyweb=\"urn:gip-hyweb-com\" xmlns=\"\">");
            sb.AppendLine("<li><a href=\"javascript:getSelectedText('ReportId="
                + ReportId
                + "&CategoryId="
                + CategoryId
                + "&ActorType="
                + ActorType
                + "')\" class=\"Rword\">推薦詞彙</a></li>");
            sb.AppendLine("<li><a href=\"#\" class=\"Rword\" title=\"系統問題\" onclick=\"window.showModalDialog('http://" + Request.Url.Authority + "/mailbox.asp?a=a9191',self);return false;\">");
            sb.AppendLine("系統問題</a></li>");
            sb.Append("<input type=\"hidden\"  name=\"type\" id=\"type\" value=\"3\" />");
            sb.Append("<input type=\"hidden\"  name=\"ARTICLE_ID\" id=\"ARTICLE_ID\" value=" + ReportId + " />");

            sb.AppendLine("<li><a class=\"Print\" href='#' onclick=\"window.open('categoryprintcontent.aspx" + Request.Url.Query + "')\" title=\"友善列印\" >");
            sb.AppendLine("友善列印</a></li>");
            sb.AppendLine("<li><a href=\"javascript:history.go(-1);\" class=\"Back\" title=\"回上一頁\">回上一頁</a><noscript>\t本網頁使用SCRIPT編碼方式執行回上一頁的動作，如果您的瀏覽器不支援SCRIPT，請直接使用「Alt+方向鍵向左按鍵」來返回上一頁</noscript></li></ul>");
            sb.AppendLine("<div id=\"cp\" xmlns=\"\">");
            sb.AppendLine("<table summary=\"排版表格\" class=\"cptable\">");
            sb.AppendLine("<col class=\"title\">");
            sb.AppendLine("<col class=\"cptablecol2\">");
            //Modify by Max 2011/06/03
            sb.AppendLine("<tr><th scope=\"row\">標題</th><td>" + PediaUtility.ReplacePedia(documentDetailInfo.VersionTitle).GetValue(0) + "</td></tr>");
            Array Summary = new Array[1];
            Summary = PediaUtility.ReplacePedia(documentDetailInfo.VersionSummary);
            sb.AppendLine("<tr><th scope=\"row\">內文</th><td>" + Summary.GetValue(0) + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">文件屬性</th><td>" + PediaUtility.ReplacePedia(documentDetailInfo.DocumentClass.ClassName).GetValue(0) + "</td></tr>");
            //End
			if (result.Elements.Count>0)
			{
            sb.AppendLine("<tr><th scope=\"row\">點閱次數</th><td>" + result.Elements[0].Clix + "</td></tr>");
			}
            sb.AppendLine("<tr><th scope=\"row\">作者</th><td>" + documentDetailInfo.DocumentAttributes[1].Value + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">知識樹分類</th><td>" + ListCategory(documentDetailInfo) + "</td></tr>");
            foreach (DocumentAttributeInfo daInfo in documentDetailInfo.DocumentAttributes)
            {
                if (daInfo.DisplayName == "公佈日期" && !string.IsNullOrEmpty(daInfo.Value))
                {
                    sb.AppendLine("<tr><th scope=\"row\">公佈日期</th><td>" +
                        DateTime.Parse(WebUtility.ParseISO8601SimpleString(daInfo.Value, DateTimeKind.Utc)).ToShortDateString()
                        + "</td></tr>");
                    break;
                }
            }

            sb.AppendLine("</table>");

            attachmentFiles = documentDetailInfo.SourceFiles;
            if (attachmentFiles.Length > 0)
            {
                sb.AppendLine("<div class=\"download\"><h5>相關檔案下載</h5>");
                foreach (SourceFileInfo sourceFileInfo in attachmentFiles)
                {
                    //api.DownloadFile(
                    //link = WebUtility.GetAppSetting("LambdaWebSite")
                    //    + @"download.aspx?documentId="
                    //    + ReportId
                    //    + @"&fileName="
                    //    + HttpUtility.UrlEncode(sourceFileInfo.DisplayName)
                    //    + @"&ver=" + documentDetailInfo.VersionNumber.ToString();

                    //sb.AppendLine("<ul><li><a target=\"_blank\" href=\""
                    //    + link
                    //    + "\" title=\"" + sourceFileInfo.DisplayName + "\">"
                    //    + sourceFileInfo.DisplayName + "</a> (" + sourceFileInfo.FileSize.ToString() + " KB) </li></ul>");

                    link = WebUtility.GetKMAPISetting("APIUrl") + "/" + @"download.aspx?documentId="
                       + ReportId
                       + @"&fileName="
                       + HttpUtility.UrlEncode(sourceFileInfo.DisplayName)
                       + @"&ver=" + documentDetailInfo.VersionNumber.ToString();

                    sb.AppendLine("<ul><li><a target=\"_blank\" href=\""
                        + link
                        + "\" title=\"" + sourceFileInfo.DisplayName + "\">"
                        + sourceFileInfo.DisplayName + "</a> (" + sourceFileInfo.FileSize.ToString() + " KB) </li></ul>");


                }
                sb.AppendLine("</div>");
            }
            sb.AppendLine("</div>");
        }
        else
        {
            WebUtility.WindowAlertAndBack(Page, "您無權限閱讀此文章");
            return;
        }

        TableText.Text = sb.ToString();

        #endregion

        #region 設定延申閱讀

        StringBuilder recommendBuilder = new StringBuilder();
        try
        {
            DocumentInfo[] recommendDocuments = api.recommend.GetDocuments(new int[] { int.Parse(ReportId) }, 100);

            if (recommendDocuments.Length > 0)
            {
                int index = 1;
                recommendBuilder.Append("<table class=\"similar\">");
                recommendBuilder.Append("<tr><th colSpan=\"3\">延伸閱讀</th></tr>");
                IList<DocumentInfo> recommendList = new List<DocumentInfo>();
                Hashtable categoryList = new Hashtable();
                Hashtable onlineDateList = new Hashtable();
                DateTime onlineDate;
                foreach (DocumentInfo documentInfo in recommendDocuments)
                {
                    if (recommendList.Count >= 5)
                    { break; }
                    DocumentDetailInfo ddi = api.documents.Get(documentInfo.DocumentId);
                    cai = ddi.Categories;
                    if (!IsKmwebCategory(cai))
                        continue;
                    DocumentAttributeInfo[] dai = ddi.DocumentAttributes;
                    bool isRead = false;
                    bool isOnline = false;
                    foreach (DocumentAttributeInfo documentAttributeInfo in dai)
                    {
                        if (documentAttributeInfo.DisplayName == "公佈日期" && !string.IsNullOrEmpty(documentAttributeInfo.Value))
                        {

                            if (DateTime.Parse(WebUtility.ParseISO8601SimpleString(documentAttributeInfo.Value, DateTimeKind.Utc)) <= DateTime.Now)
                            {
                                onlineDateList.Add(documentInfo.DocumentId,
                                    DateTime.Parse(WebUtility.ParseISO8601SimpleString(documentAttributeInfo.Value, DateTimeKind.Utc)).ToShortDateString());
                                isOnline = true;
                            }
                        }

                        if (documentAttributeInfo.DisplayName == "可閱讀分眾導覽(前端入口網)"
                            && documentAttributeInfo.Value.Contains(WebUtility.ConvertGroupsReadingInInter(ActorType)))
                        { isRead = true; }

                        if (isRead && isOnline)
                        { break; }
                    }

                    if (isRead && isOnline)
                    {
                        recommendList.Add(documentInfo);

                        foreach (CategoryInfo ci in ddi.Categories)
                        {
                            CategoryInfoCollection categoryInfos = ci.Path;
                            bool isNew = false;
                            foreach (CategoryInfo categoryInfo in categoryInfos)
                            {
                                if (categoryInfo.CategoryId == int.Parse(WebUtility.GetAppSetting("CategoryTreeRootNode")))
                                {
                                    isNew = true;
                                    break;
                                }
                            }
                            if (isNew)
                            {
                                categoryList.Add(documentInfo.DocumentId, ci);
                                break;
                            }
                        }
                    }
                }

                foreach (DocumentInfo documentInfo in recommendList)
                {
                    recommendBuilder.Append("<tr><td width=\"6%\">" + index + "</td>");
                    recommendBuilder.Append("<td width=\"74%\"><a href=\"/category/categorycontent.aspx?ReportId="
                        + documentInfo.DocumentId.ToString()
                        + "&CategoryId="
                        + (categoryList.ContainsKey(documentInfo.DocumentId) ?
                        ((CategoryInfo)categoryList[documentInfo.DocumentId]).CategoryId.ToString() : CategoryId)
                        + "&ActorType="
                        + ActorType
                        + "\">"
                        + documentInfo.VersionTitle
                        + "</a></td>");
                    recommendBuilder.Append("<td width=\"20%\">["
                        + onlineDateList[documentInfo.DocumentId].ToString()
                        + "]</td></tr>");
                    index++;
                }
                recommendBuilder.Append("</table><BR/>");
            }
        }
        catch (Exception ex)
        {
            Response.Write(ex.ToString());
            Response.End();
        }
        relationsArticle = recommendBuilder.ToString();

        #endregion

        #region KPI給分
        if (Session["memID"] != null && !string.IsNullOrEmpty(Session["memID"].ToString()))
        {
            int count = 0;
            if (attachmentFiles == null || attachmentFiles.Length <= 0)
            { count++; }
            else
            { count = attachmentFiles.Length; }
            KPIBrowse browse = new KPIBrowse(Session["memID"].ToString(),
                "browseCatTreeCP",
                count.ToString());
            browse.HandleBrowse();
        }
        #endregion

        #region 統計文章瀏覽數
        using (SqlConnection cn = new SqlConnection())
        {
            string connString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["LambdaCoaConnString"].ConnectionString;

            if (null != connString)
            {
                cn.ConnectionString = connString;
                cn.Open();

                SqlCommand cmd = new SqlCommand("SP_PerformanceStatisticsADD_DG", cn);
                cmd.CommandType = CommandType.StoredProcedure;

                cmd.Parameters.Add("@docId", SqlDbType.Int);
                cmd.Parameters.Add("@D", SqlDbType.Bit);
                cmd.Parameters.Add("@G", SqlDbType.Bit);
                cmd.Parameters["@docId"].Value = Convert.ToInt32(ReportId);
                cmd.Parameters["@D"].Value = 1;
                cmd.Parameters["@G"].Value = 0;

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw ex.GetBaseException();
                }
                finally
                {
                    cn.Close();
                }
            }
        }
        try
        {
            //使用 WebClient 類別向伺服器發出要求
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            string serviceUrl = WebUtility.GetKMAPISetting("APIUrl") + "/hit/document/create/" + ReportId + "?" + WebUtility.GetAPIParameterString();
            //Response.Write(serviceUrl);
            //反序列化用
            string resultData = "";
            //接收結果
            resultData = client.DownloadString(serviceUrl);
        }
        catch { }
        #endregion
    }
    #endregion

    #region 計算是否是農業知識樹
    private bool IsKmwebCategory(CategoryInfo[] ci)
    {
        foreach (CategoryInfo ca in ci)
        {
            if (ca.CategoryId != 0)
            {
                CategoryInfo[] paths = api.categories.GetPath(ca.CategoryId);
                if (paths.Length > 0)
                {
                    foreach (CategoryInfo categoryInfo in paths)
                    {
                        if (categoryInfo.CategoryId.ToString() == WebUtility.GetAppSetting("CategoryTreeRootNode"))
                        {
                            return true;
                        }
                    }
                }
            }
        }
        return false;
    }
    #endregion

    #region Methods

    /// <summary>
    /// 列出分類清單
    /// </summary>
    /// <param name="documentDetailInfo">DocumentDetailInfo</param>
    /// <returns>String</returns>
    private string ListCategory(DocumentDetailInfo documentDetailInfo)
    {
        StringBuilder result = new StringBuilder();

        if (documentDetailInfo.Categories != null && documentDetailInfo.Categories.Length > 0)
        {
            string s = string.Empty;
            for (int i = 0; i < documentDetailInfo.Categories.Length; i++)
            {
                CategoryInfoCollection categoryInfos = documentDetailInfo.Categories[i].Path;
                s = NavTitleText.Text;
                bool isNew = false;
                foreach (CategoryInfo categoryInfo in categoryInfos)
                {
                    if (categoryInfo.DisplayName != "#root#" && categoryInfo.CategoryId.ToString()
                        != WebUtility.GetAppSetting("CategoryTreeRootNode"))
                    { s += " > " + categoryInfo.DisplayName; }
                    if (categoryInfo.CategoryId == int.Parse(WebUtility.GetAppSetting("CategoryTreeRootNode")))
                    { isNew = true; }
                }
                if (isNew)
                { result.Append(s + "<br>"); }
            }
        }
        return result.ToString();
    }

    #endregion

}
