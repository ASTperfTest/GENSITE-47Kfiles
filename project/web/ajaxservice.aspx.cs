#region Imports
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.API.Client;
using System.Collections.Specialized;
using GSS.Vitals.API.Client.Contracts.Index;
using System.Text;
using GSS.Vitals.API.Client.Contracts.KnowledgeBase;
using Newtonsoft.Json;
using Jayrock.Json;
using Newtonsoft.Json.Linq;


#endregion

public partial class ajaxservice : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string command = WebUtility.GetStringParameter("cmd");
        JObject jsonObject = JsonConvert.DeserializeObject<JObject>(command);

        processCommand(jsonObject["id"].ToString());
    }

    private void processCommand(string command)
    {
        switch (command)
        {
            case "\"gethotdocuments\"":
                doGetHotDocuments();
                break;
        }
    }

    private void doGetHotDocuments()
    {
        try
        {
            StringBuilder sb = new StringBuilder();
            JObject jsonObject = JsonConvert.DeserializeObject<JObject>(WebUtility.GetStringParameter("cmd"));
            string key = jsonObject["key"].ToString().ToUpper();
            string cachekey = "ajaxServiceCacheKey" + key;
                        

            if (Cache[cachekey] == null)
            {
                #region 取得輸出字串

                RestRpcClient apiClient = WebUtility.GetAPIClient(WebUtility.GetKMAPISetting("APIUrl"),
                    WebUtility.GetKMAPISetting("APIKey"),
                   WebUtility.GetKMAPISetting("APIActor"),
                   WebUtility.GetKMAPISetting("APIFormat"),
                   10);

                //#用進階查詢取得文章
                NameValueCollection keyValues = new NameValueCollection();
                keyValues.Add("公佈日期", "[19000101160000 TO " + DateTime.UtcNow.ToString("yyyyMMddHHmmss") + "]");
                keyValues.Add("可閱讀分眾導覽\\(前端入口網\\)", "A,B");
                ExtendedSearchResultPagedCollection result = apiClient.search.ByAdvanced
                    (null, null, null, new int[] { 2 },
                    null, string.Empty, null, null, string.Empty, keyValues, false, true, false, false,
                    (key == "\"A\"" ? SpecialFields.FieldType.CreationDatetime : SpecialFields.FieldType.Clix));

                string link = string.Empty;
                if (result.Elements != null && result.Elements.Count > 0)
                {
                    ExtendedSearchResult searchResult;

                    for (int i = 0; i <= result.Elements.Count - 1; i++)
                    {
                        searchResult = result.Elements[i];
                        DocumentDetailInfo documentDetailInfo = apiClient.documents.Get(int.Parse(searchResult.UniqueKey));
                        bool isOnline = false;
                        DocumentAttributeInfo documentAttributeInfo = new DocumentAttributeInfo();
                        foreach (DocumentAttributeInfo daInfo in documentDetailInfo.DocumentAttributes)
                        {
                            if (daInfo.DisplayName == "公佈日期" && !string.IsNullOrEmpty(daInfo.Value))
                            {
                                if (DateTime.Parse(WebUtility.ParseISO8601SimpleString(daInfo.Value, DateTimeKind.Utc)) <= DateTime.Now)
                                {
                                    isOnline = true;
                                    documentAttributeInfo = daInfo;
                                }
                                break;
                            }
                        }
                        if (isOnline)
                        {
                            sb.Append("<li>");
                            string CategoryId = string.Empty;
                            CategoryId = searchResult.Categories[searchResult.Categories.Length - 1].ToString();
                            sb.Append("<a href=\"category/categorycontent.aspx?ReportId=" + searchResult.UniqueKey);
                            sb.Append("&CategoryId=" + CategoryId);
                            CategoryInfo categoryInfo = apiClient.categories.Get(int.Parse(CategoryId), false);
                            sb.Append("&ActorType=002\">" + searchResult.Title + "</a>");
                            //#分類,文件類型,點閱數,公佈日期
                            sb.Append("[" + categoryInfo.DisplayName + "]");
                            sb.Append("[" + ((DocumentClassInfo)WebUtility.GetBuildDocumentClassList(Page)[searchResult.DocumentClass]).ClassName + "]");
                            sb.Append("[" + searchResult.Clix.ToString() + "]");
                            sb.Append("[" + searchResult.CreationDatetime.ToString("yyyy/MM/dd")+ "]");
                            if (key == "B")
                            { sb.Append(" (" + searchResult.Clix + ")"); }
                        }
                    }

                    if (sb.Length > 0)
                    {
                        Cache.Add(cachekey, sb, null
                        , DateTime.Now.AddHours(1)
                        , System.Web.Caching.Cache.NoSlidingExpiration
                        , System.Web.Caching.CacheItemPriority.Default, null);
                    }
                }
                else
                {
                    sb.Append("查無資料！");
                }

                #endregion
            }
            else
            {
                sb = (StringBuilder)Cache[cachekey];
            }            

            Response.Write(sb.ToString());
        }
        catch (Exception ex)
        {
            Response.Write(ex.Message);
        }

    }
}
