#region Imports
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.API.Client;
using GSS.Vitals.API.Client.Contracts.KnowledgeBase;
using System.Text;
using GSS.Vitals.API.Client.Contracts.Index;
using System.Collections;
#endregion

public partial class Category_categoryprintcontent : System.Web.UI.Page
{
    #region Fields
    private RestRpcClient api = null;
    protected string relationsArticle = string.Empty;
    #endregion

    #region Page Event
    protected void Page_Load(object sender, EventArgs e)
    {
        string xssTest = Request.RawUrl;
        if (WebUtility.checkParam(xssTest))
        {
            Response.Write("<script>alert('網址中包含不正常的參數,頁面將導回首頁!!');window.location.href='/mp.asp';</script>");
            Response.End();
        }

        string CategoryId = WebUtility.GetStringParameter("CategoryId");
        string ActorType = WebUtility.GetStringParameter("ActorType");
        string ReportId = WebUtility.GetStringParameter("ReportId");
        api = WebUtility.GetAPIClient();

        string link = "";
        link = "categorylist.aspx?CategoryId={catid}&ActorType={actortype}";

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
        DocumentDetailInfo documentDetailInfo = api.documents.Get(int.Parse(ReportId));
        ExtendedSearchResultPagedCollection result = api.search.ByAdvanced("_l_unique_key:" + ReportId,
                    null, null, null,
                    null, string.Empty, null, null, string.Empty, null, false, false, false, false, SpecialFields.FieldType.Clix);
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

            sb.AppendLine("<div id=\"cp\" xmlns=\"\">");
            sb.AppendLine("<table summary=\"排版表格\" class=\"cptable\">");
            sb.AppendLine("<col class=\"title\">");
            sb.AppendLine("<col class=\"cptablecol2\">");
            sb.AppendLine("<tr><th scope=\"row\">標題</th><td>" + documentDetailInfo.VersionTitle + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">內文</th><td>" + documentDetailInfo.VersionSummary + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">文件屬性</th><td>" + documentDetailInfo.DocumentClass.ClassName + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">點閱次數</th><td>" + result.Elements[0].Clix + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">作者</th><td>" + documentDetailInfo.DocumentAttributes[1].Value + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">知識樹分類</th><td>" + ListCategory(documentDetailInfo) + "</td></tr>");
            sb.AppendLine("<tr><th scope=\"row\">張貼日期</th><td>" + documentDetailInfo.CreationDatetime.ToShortDateString() + "</td></tr>");
            sb.AppendLine("</table>");

            attachmentFiles = documentDetailInfo.SourceFiles;
            if (attachmentFiles.Length > 0)
            {
                sb.AppendLine("<div class=\"download\"><h5>相關檔案下載</h5>");
                foreach (SourceFileInfo sourceFileInfo in attachmentFiles)
                {
                    link = WebUtility.GetAppSetting("LambdaWebSite")
                        + @"download.aspx?documentId="
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
                int size = 0;
                foreach (DocumentInfo documentInfo in recommendDocuments)
                {
                    if (size >= 5)
                    { break; }

                    DocumentDetailInfo ddi = api.documents.Get(documentInfo.DocumentId);
                    DocumentAttributeInfo[] dai = ddi.DocumentAttributes;
                    bool isRead = true;
                    foreach (DocumentAttributeInfo documentAttributeInfo in dai)
                    {
                        if (documentAttributeInfo.DisplayName == "公布日期" && !string.IsNullOrEmpty(documentAttributeInfo.Value))
                        {
                            DateTime onlineDate;
                            bool canPase = DateTime.TryParse(documentAttributeInfo.Value, out onlineDate);
                            if (canPase && onlineDate >= DateTime.Now)
                            {
                                isRead = false;
                                break;
                            }
                        }
                        if (documentAttributeInfo.DisplayName == "可閱讀分眾導覽(前端入口網)"
                            && !documentAttributeInfo.Value.Contains(WebUtility.ConvertGroupsReadingInInter(ActorType)))
                        {
                            isRead = false;
                            break;
                        }
                    }

                    if (isRead)
                    {
                        recommendList.Add(documentInfo);
                        size++;
                    }
                }

                foreach (DocumentInfo documentInfo in recommendList)
                {
                    recommendBuilder.Append("<tr><td width=\"6%\">" + index + "</td>");
                    recommendBuilder.Append("<td width=\"74%\"><a href=\"/category/categorycontent.aspx?ReportId="
                        + documentInfo.DocumentId.ToString()
                        + "&CategoryId="
                        + CategoryId
                        + "&ActorType="
                        + ActorType
                        + "\">"
                        + documentInfo.VersionTitle
                        + "</a></td>");
                    recommendBuilder.Append("<td width=\"20%\">["
                        + documentInfo.CreationDatetime.ToShortDateString()
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
    }
    #endregion

    #region Methods
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
