#region Imports
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.API.Client;
using GSS.Vitals.API.Client.Contracts.KnowledgeBase;
using System.Xml;
using System.Collections;
using System.Text;
using GSS.Vitals.API.Client.Contracts.Report;
using GSS.Vitals.API.Client.Contracts.Index;
using System.Collections.Specialized;
using System.Data;
#endregion

public partial class Category_categorylist : System.Web.UI.Page
{

    #region Fields
    private RestRpcClient api = null;
    private string UrlStr = string.Empty;
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
        string ActorId = string.Empty;

        CatTreeViewDS.Data = "";
        api = WebUtility.GetAPIClient();

        #region Exception
        if (Request.RawUrl.Contains("%3c") ||
            Request.RawUrl.Contains("%3d") ||
            Request.RawUrl.Contains("%3e") ||
            Request.RawUrl.Contains("%22") ||
            Request.RawUrl.Contains("%"))
        {
            Response.Redirect("/mp.asp?mp=1");
            Response.End();
        }
        if (!string.IsNullOrEmpty(CategoryId))
        {
            int parseInt;
            if (!int.TryParse(CategoryId, out parseInt))
            {
                Response.Redirect("/mp.asp?mp=1");
                Response.End();
            }
        }
        if (Session["gstyle"] == null)
        {
            ActorId = "";
        }
        else
        {
            ActorId = Convert.ToString(Session["gstyle"]);
        }

        //---defalut is 知識庫首頁---
        if (string.IsNullOrEmpty(ActorType))
        {
            ActorType = "000";
        }

        if (ActorType != "000" & ActorType != "001" & ActorType != "002" & ActorType != "003")
        {
            Response.Redirect("/mp.asp?mp=1");
            Response.End();
        }
        //---學者---
        if (ActorType == "003")
        {
            //---檢查---
            if (ActorId == "003" | ActorId == "005")
            {
                //---可閱讀---
            }
            else
            {
                //---無權限---
                WebUtility.WindowAlertAndBack(Page, "限學者會員方可進入!");
                return;
            }
        }
        #endregion

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

        if (string.IsNullOrEmpty(CategoryId))
        { CategoryId = WebUtility.GetAppSetting("CategoryTreeRootNode"); }
        XmlDocument doc = new XmlDocument();
        ArrayList data = new ArrayList();
        string rootDisplayName = string.Empty;
        if (ActorType == "000" || ActorType == "002")
        { rootDisplayName = "消費者知識庫"; }
        else
        { rootDisplayName = NavTitleText.Text; }
        WebUtility.GenerateXMLDataSource(doc, int.Parse(CategoryId), rootDisplayName, UrlStr, data);

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

        #region 設定子目錄資料數
        StringBuilder sbb = new StringBuilder();
        string defaulturl = "categorylist.aspx?CategoryId={0}&ActorType={1}&t=a";
        string tempurl = string.Format(defaulturl, CategoryId, ActorType);
        string tempnode = "<a href=\"{0}\">{1}({2})</a>｜";

        bool isScholar = false;
        if (Session["gstyle"] != null && !string.IsNullOrEmpty(Session["gstyle"].ToString()))
        {
            if (Session["gstyle"].ToString() == "003" || Session["gstyle"].ToString() == "005")
            { isScholar = true; }
        }
        ExtendReportData extenReportData = new ExtendReportData();
        if (ActorType != "000")
        {
            sbb.Append("<div id=dividetype>子目錄資料數：");

            if (ActorType == "003")
            { extenReportData = api.report.GetReportByCategoryIdAndCondition(int.Parse(CategoryId), (isScholar ? WebUtility.ConvertGroupsReadingInInter(ActorType) : "")); }
            else
            {
                try
                {
                    extenReportData = api.report.GetReportByCategoryIdAndCondition(int.Parse(CategoryId),
                        WebUtility.ConvertGroupsReadingInInter(ActorType));
                }
                catch (InvalidOperationException ex)
                {
                    Response.Redirect("/mp.asp?mp=1");
                }

            }

            if (extenReportData != null)
            {
                for (int index = 1; index < extenReportData.Data.Count; index++)
                {
                    sbb.Append(tempnode.Replace("{0}",
                        string.Format(defaulturl, extenReportData.Data[index][0].ToString(), ActorType)).Replace("{1}"
                        , extenReportData.Data[index][1].ToString()).Replace("{2}",
                         extenReportData.Data[index][2].ToString()));
                }
            }

            sbb.Append("</div>");
        }
        else if (!string.IsNullOrEmpty(WebUtility.GetStringParameter("t", string.Empty)))
        {
            string groupDisplayName = "「消費者」、「生產者」、「學者」";
            sbb.Append("<div class=nodata>請選擇上方" + groupDisplayName + "分眾知識樹開始瀏覽，其中學者知識樹需具有學者會員資格。選擇分眾知識樹後，您可直接點選左欄欲瀏覽之知識庫節點。</div>");
            sbb.Append("<div class=hotessay><h3>熱門知識庫文章</h3><div id=\"MagTabs\"><ul>");

            if (WebUtility.GetStringParameter("t") == "a")
            {
                sbb.Append("<li class=current><a href=\"" + tempurl.Replace("{0}", "") + "\"><span>最新文章</span></a></li>");
                sbb.Append("<li><a href=\"" + tempurl.Replace("&t=a", "&t=b").Replace("{0}", "") + "\"><span>最多瀏覽</span></a></li>");
            }
            else
            {
                sbb.Append("<li><a href=\"" + tempurl.Replace("{0}", "") + "\"><span>最新文章</span></a></li>");
                sbb.Append("<li class=current><a href=\"" + tempurl.Replace("&t=a", "&t=b").Replace("{0}", "") + "\"><span>最多瀏覽</span></a></li>");
            }

            sbb.Append("</ul></div></div>");
        }

        NodeText.Text = sbb.ToString();

        #endregion

        #region 列出目前分類下的文章
        int PageSize = 10;
        int PageNumber = 0;
        string defaultMessage = @"<div class=nodata>本層節點目前無資料，
                             您可點選上方子目錄資料項目、或點選左欄其他知識庫節點。</div>";
        if (!Page.IsPostBack)
        {
            try
            {
                if (string.IsNullOrEmpty(WebUtility.GetStringParameter("PageSize")))
                {
                    PageSize = 10;
                }
                else
                {
                    PageSize = Convert.ToInt32(WebUtility.GetStringParameter("PageSize"));
                }
                if (string.IsNullOrEmpty(WebUtility.GetStringParameter("PageNumber")))
                {
                    PageNumber = 1;
                }
                else
                {
                    PageNumber = Convert.ToInt32(WebUtility.GetStringParameter("PageNumber"));
                }
            }
            catch (Exception ex)
            {
                //Response.Write(ex.Message);
                //return;
                Response.Redirect("/mp.asp?mp=1");
            }
        }
        else
        {
            PageSize = int.Parse(PageSizeDDL.SelectedValue);
            PageNumber = int.Parse(PageNumberDDL.SelectedValue);
        }

        int total = 0;
        int pageCount = 0;
        int position = 1;

        StringBuilder sb = new StringBuilder();        
        if (!string.IsNullOrEmpty(CategoryId))
        {
            try
            {
                //#用進階查詢取得文章
                NameValueCollection keyValues = new NameValueCollection();
                keyValues.Add("公佈日期", "[19000101160000 TO " + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + "]");
                bool isContainChild = false;
                if (ActorType == "000" && CategoryId == WebUtility.GetAppSetting("CategoryTreeRootNode")
                && !string.IsNullOrEmpty(WebUtility.GetStringParameter("t", string.Empty)))
                { keyValues.Add("可閱讀分眾導覽\\(前端入口網\\)", "A,B"); isContainChild = true; }
                else
                {
                    keyValues.Add("可閱讀分眾導覽\\(前端入口網\\)", (ActorType == "003" ? (isScholar ? WebUtility.ConvertGroupsReadingInInter(ActorType) : "")
                       : (isScholar ? WebUtility.ConvertGroupsReadingInInter(ActorType) + ",C" : WebUtility.ConvertGroupsReadingInInter(ActorType))));
                }
                
                ExtendedSearchResultPagedCollection result;
                string cachekey = "ExtendedSearchResultPagedCollection_" + Request.Url.Query;
                //Response.Write(Request.Url.Query);
                if (Cache[cachekey] == null)
                {
                    result = api.search.ByAdvanced 
                        (null, null, null, new int[] { int.Parse(CategoryId) },
                        null, string.Empty, null, null, string.Empty, keyValues, false, isContainChild, false, false,
                        (WebUtility.GetStringParameter("t", string.Empty) == string.Empty ? SpecialFields.FieldType.CreationDatetime : (WebUtility.GetStringParameter("t") == "a" ? SpecialFields.FieldType.CreationDatetime : SpecialFields.FieldType.Clix)));
                    if (result != null)
                    {
                        if (result.Elements.Count > 0)
                        {
                            Cache.Add(cachekey, result, null
                                , DateTime.Now.AddHours(1)
                                , System.Web.Caching.Cache.NoSlidingExpiration
                                , System.Web.Caching.CacheItemPriority.Default, null);
                        }
                    }
                }
                else
                {
                    result = (ExtendedSearchResultPagedCollection)Cache[cachekey]; ;
                }


                if (result.Elements != null && result.Elements.Count > 0)
                {
                    total = result.Elements.Count;
                    pageCount = int.Parse(Convert.ToString(total / PageSize + 1).ToString());

                        #region 設定頁數
                        if (pageCount < PageNumber)
                        { PageNumber = pageCount; }

                        PageNumberText.Text = PageNumber.ToString();
                        TotalPageText.Text = pageCount.ToString();
                        TotalRecordText.Text = total.ToString();
                        if (PageSize == 10)
                        {
                            PageSizeDDL.SelectedIndex = 0;
                        }
                        else if (PageSize == 20)
                        {
                            PageSizeDDL.SelectedIndex = 1;
                        }
                        else if (PageSize == 30)
                        {
                            PageSizeDDL.SelectedIndex = 2;
                        }
                        else if (PageSize == 50)
                        {
                            PageSizeDDL.SelectedIndex = 3;
                        }

                        ListItem item = default(ListItem);
                        int j = 0;
                        PageNumberDDL.Items.Clear();
                        for (j = 0; j <= pageCount - 1; j++)
                        {
                            item = new ListItem();
                            item.Value = Convert.ToString(j + 1);
                            item.Text = Convert.ToString(j + 1);
                            if (PageNumber == (j + 1))
                            {
                                item.Selected = true;
                            }
                            PageNumberDDL.Items.Insert(j, item);
                            item = null;
                        }
                        #endregion

                        #region 設定上下頁Link
                        string PrevLinkUrl = "categorylist.aspx?CategoryId="
                                + (ActorType != "000" ? CategoryId : "")
                                + (ActorType != "000" ? "" : "&t=" + WebUtility.GetStringParameter("t"))
                                + "&ActorType="
                                + ActorType
                                + "&PageNumber="
                                + Convert.ToString(PageNumber - 1)
                                + "&PageSize="
                                + PageSize;

                        string NextLinkUrl = "categorylist.aspx?CategoryId="
                                + (ActorType != "000" ? CategoryId : "")
                                + (ActorType != "000" ? "" : "&t=" + WebUtility.GetStringParameter("t"))
                                + "&ActorType="
                                + ActorType
                                + "&PageNumber="
                                + Convert.ToString(PageNumber + 1)
                                + "&PageSize="
                                + PageSize;

                        if (PageNumber > 1)
                        {
                            PreviousText.Visible = true;
                            PreviousImg.Visible = true;
                            PreviousLink.NavigateUrl = PrevLinkUrl;
                        }
                        else
                        {
                            PreviousText.Visible = false;
                            PreviousImg.Visible = false;
                        }
                        if (Convert.ToInt32(PageNumberDDL.SelectedValue) < pageCount)
                        {
                            NextText.Visible = true;
                            NextImg.Visible = true;
                            NextLink.NavigateUrl = NextLinkUrl;
                        }
                        else
                        {
                            NextText.Visible = false;
                            NextImg.Visible = false;
                        }
                        #endregion

                    //#設定初始位置
                    position = PageSize * (PageNumber - 1);
                    int i = 0;
                    link = "";
                    Hashtable dcList = WebUtility.GetBuildDocumentClassList(Page);
                    ExtendedSearchResult searchResult;

                    #region 設定目前分類下文章內容
                    switch (ActorType)
                    {
                        case "000":
                            sb.Append("<ul class=list>");

                            for (i = 0; i <= PageSize - 1; i++)
                            {
                                
                                searchResult = result.Elements[position];

                                sb.Append("<li>");

                                sb.Append("<a href=\"categorycontent.aspx?ReportId=" + searchResult.UniqueKey);

                                sb.Append("&CategoryId=" + CategoryId);
                                CategoryInfo categoryInfo = api.categories.Get(int.Parse(CategoryId), false);
                                sb.Append("&ActorType=002\">" + searchResult.Title + "</a>");

                                if (!string.IsNullOrEmpty(WebUtility.GetStringParameter("t", string.Empty)))
                                { sb.Append("[" + categoryInfo.DisplayName + "]"); }
                                sb.Append("[" + ((DocumentClassInfo)WebUtility.GetBuildDocumentClassList(Page)[searchResult.DocumentClass]).ClassName + "]");
                                sb.Append("[" + searchResult.Clix + "]");
                                // Grace改為文件建立日期 故將504行註解
                                string[] stringArray = searchResult.CreationDatetime.ToString().Split(' ');
                                sb.Append("[" + stringArray[0] + "]");

                                //bob快取文件的公佈日期
                                string DocPublishedDateCacheKey = "DDI_" + searchResult.UniqueKey;                                
                                if (Cache[DocPublishedDateCacheKey] == null)
                                {
                                    Cache.Add(DocPublishedDateCacheKey, "", null
                                        , DateTime.Now.AddHours(6)
                                        , TimeSpan.Zero
                                        , System.Web.Caching.CacheItemPriority.Default
                                        , null);

                                    DocumentDetailInfo documentDetailInfo = documentDetailInfo = api.documents.Get(int.Parse(searchResult.UniqueKey));
                                    //DocumentAttributeInfo documentAttributeInfo = new DocumentAttributeInfo();
                                    foreach (DocumentAttributeInfo daInfo in documentDetailInfo.DocumentAttributes)
                                    {
                                        try
                                        {
                                            if (daInfo.DisplayName == "公佈日期" && !string.IsNullOrEmpty(daInfo.Value))
                                            {
                                                string pDate = string.Format("[{0}]"
                                                    , DateTime.Parse(WebUtility.ParseISO8601SimpleString(daInfo.Value, DateTimeKind.Utc)).ToShortDateString());

                                                Cache[DocPublishedDateCacheKey] = pDate ?? string.Empty;
                                                break;
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                //sb.Append(Cache[DocPublishedDateCacheKey]);
                                if (string.IsNullOrEmpty(Cache[DocPublishedDateCacheKey].ToString())) 
                                    Cache.Remove(DocPublishedDateCacheKey);
                                //==================================================================

                                sb.Append("</li>");
                                position += 1;
                                if (result.Elements.Count <= position)
                                { break; }
                                searchResult = result.Elements[position];
                            }
                            sb.Append("</ul>");
                            sb.Append("</div>");
                            break;
                        default:
                            sb.Append("<div class=\"list\">");
                            sb.Append("<ul>");

                            for (i = 0; i <= PageSize - 1; i++)
                            {
                                link = "";
                                sb.Append("<li>");
                                searchResult = result.Elements[position];

                                link = "<a href=\"categorycontent.aspx?ReportId="
                                    + searchResult.UniqueKey;

                                link += "&CategoryId="
                                    + CategoryId
                                    + "&ActorType="
                                    + ActorType + "\">" + searchResult.Title + "</a>";

                                sb.Append("<span>" + link + "</span>");
                                sb.Append("<span>" + searchResult.Clix + "</span>");
                                sb.Append("<span>" + ((DocumentClassInfo)dcList[searchResult.DocumentClass]).ClassName + "</span>");

                                //bob快取文件的公佈日期
                                string DocPublishedDateCacheKey = "DDI_" + searchResult.UniqueKey;
                                if (Cache[DocPublishedDateCacheKey] == null)
                                {
                                    Cache.Add(DocPublishedDateCacheKey, "", null
                                        , DateTime.Now.AddHours(6)
                                        , TimeSpan.Zero
                                        , System.Web.Caching.CacheItemPriority.Default
                                        , null);

                                    DocumentDetailInfo documentDetailInfo = documentDetailInfo = api.documents.Get(int.Parse(searchResult.UniqueKey));
                                    //DocumentAttributeInfo documentAttributeInfo = new DocumentAttributeInfo();
                                    foreach (DocumentAttributeInfo daInfo in documentDetailInfo.DocumentAttributes)
                                    {
                                        try
                                        {
                                            if (daInfo.DisplayName == "公佈日期" && !string.IsNullOrEmpty(daInfo.Value))
                                            {
                                                string pDate = string.Format("[{0}]"
                                                    , DateTime.Parse(WebUtility.ParseISO8601SimpleString(daInfo.Value, DateTimeKind.Utc)).ToShortDateString());

                                                Cache[DocPublishedDateCacheKey] = pDate ?? string.Empty;
                                                break;
                                            }
                                        }
                                        catch { }
                                    }
                                }
                                sb.Append(Cache[DocPublishedDateCacheKey]);
                                if (string.IsNullOrEmpty(Cache[DocPublishedDateCacheKey].ToString())) Cache.Remove(DocPublishedDateCacheKey);
                                //==================================================================

                                sb.Append("</li>");
                                position += 1;
                                if (result.Elements.Count <= position)
                                {
                                    break;
                                }
                                searchResult = result.Elements[position];
                            }

                            sb.Append("</ul>");
                            sb.Append("</div>");
                            break;

                    }
                    TableText.Text = sb.ToString();
                    #endregion
                }
                else
                { TableText.Text = defaultMessage; }
            }
            catch (Exception ex)
            {

                //Response.Write(ex.ToString());
                Response.Redirect("/mp.asp?mp=1");
                Response.End();
            }
        }
        else
        {
            PageNumberDDL.SelectedIndex = 0;
            PageSizeDDL.SelectedIndex = 0;
            TableText.Text = defaultMessage;
        }

        #endregion

    }
    #endregion

    #region Methods

    /// <summary>
    /// 以分類ID取得所有所含的分類ID
    /// </summary>
    /// <param name="categoryId">分類ID</param>
    /// <returns>分類IDs</returns>
    private string GetContainCategoryByCategoryId(string categoryId)
    {
        CategoryInfoPagedCollection al = api.categories.GetChildren(Convert.ToInt32(categoryId), false);

        if (al == null)
        { return categoryId; }
        else
        {
            string childCategories = string.Empty;
            foreach (CategoryInfo categoryInfo in al.Elements)
            { childCategories += "@" + GetContainCategoryByCategoryId(Convert.ToString(categoryInfo.CategoryId)); }
            return categoryId + childCategories;
        }
    }

    /// <summary>
    /// 取得目前分類節點下所有子節點
    /// </summary>
    /// <param name="nodes">分類節點</param>
    /// <param name="key">目前分類節點</param>
    /// <returns>ArrayList</returns>
    private ArrayList GetAllContainedNodes(IList nodes, string key)
    {
        ArrayList allNodes = new ArrayList();
        allNodes.AddRange(nodes);

        foreach (int n in nodes)
        {
            string[] children = null;
            children = GetContainCategoryByCategoryId(n.ToString()).Split('@');

            if (children != null)
            {
                foreach (string id in children)
                {
                    int nodeId = Convert.ToInt32(id);
                    if (!allNodes.Contains(nodeId))
                    {
                        allNodes.Add(nodeId);
                    }
                }
            }
        }
        allNodes.TrimToSize();
        return allNodes;
    }

    private int[] GetIncludingNodes(int[] nodeIds, string resourceType)
    {

        ArrayList allNodes = null;
        allNodes = GetAllContainedNodes(nodeIds, resourceType);
        nodeIds = (int[])allNodes.ToArray(typeof(int));

        return nodeIds;
    }

    private int[] doNodeReport(string categoryID)
    {
        string key = string.Empty;
        int[] nodes = getIdsFromPost(categoryID);

        //#一律取得底下分類夾
        nodes = GetIncludingNodes(nodes, key);
        Report rpt = new Report();

        return nodes;
    }

    private int[] getIdsFromPost(string keyName)
    {
        string ids = keyName;//WebUtility.GetStringParameter(keyName);
        if (ids == null || ids.Equals(string.Empty))
        { ids = WebUtility.GetAppSetting("CategoryTreeRootNode"); }
        int[] nodes = null;
        if (ids.IndexOf("]") > 1)
        {
            nodes = WebUtility.DeserializeAjaxResult(WebUtility.GetStringParameter(keyName),
                typeof(int[])) as int[];
        }
        else
        { nodes = new int[1] { Convert.ToInt32(ids) }; }
        return nodes;
    }

    #endregion

}
