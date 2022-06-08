using System.Configuration;
using System.Web.UI;
using Jayrock.Json;
using Newtonsoft.Json;
using System.Web;
using System.Collections.Specialized;
using GSS.Vitals.API.Client;
using System;
using System.Collections;
using GSS.Vitals.API.Client.Contracts.KnowledgeBase;
using System.Xml;

/// <summary>
/// Summary description for Utility
/// </summary>
public class WebUtility
{
    public const string managementMasterPageFile = "management.master";

    public static string GetAppSetting(string key)
    {
        return ConfigurationManager.AppSettings[key];
    }

    public static void WindowAlert(Page objPage, string MsgText)
    {
        string script = "";
        script = "<script language=Javascript>alert('" + MsgText + "');</script>";
        objPage.ClientScript.RegisterStartupScript(typeof(Page), "AlertMessage", script);
    }

    public static bool CheckLogin(object id)
    {
        if (id == null || id.ToString().Trim() == string.Empty)
        {
            return false;
        }
        else
        {
            return true;
        }
    }

    public static bool IsAdmin(string id)
    {
        string admins = "|" + GetAppSetting("Admin") + "|";
        if (admins.Contains("|" + id + "|"))
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    public static string GetStringParameter(string paraName)
    {
        return GetStringParameter(paraName, string.Empty);
    }

    public static string GetStringParameter(string paraName, string defaultValue)
    {
        if (HttpContext.Current.Request.QueryString[paraName] != null && HttpContext.Current.Request.QueryString[paraName] != "null")
        {
            defaultValue = HttpContext.Current.Request.QueryString[paraName].Trim();
        }
        else
        {
            if (HttpContext.Current.Request.Form[paraName] != null)
            {
                defaultValue = HttpContext.Current.Request.Form[paraName].Trim();
            }
        }
        return defaultValue;
    }

    private static AjaxResult GenAjaxResult(bool success, string message, object data)
    {
        AjaxResult result = new AjaxResult();
        result.Success = success;
        result.Message = message;
        result.Data = data;
        return result;
    }

    public static void WriteAjaxResult(bool success, string messageKey, object data)
    {
        WriteAjaxResult(GenAjaxResult(success, messageKey, data));
    }

    public static void WriteAjaxResult(AjaxResult result)
    {
        HttpContext.Current.Response.Expires = 0;
        HttpContext.Current.Response.CacheControl = "no-cache";
        HttpContext.Current.Response.AddHeader("Pragma", "no-cache");
        HttpContext.Current.Response.Write(JsonConvert.SerializeObject(result));
    }

    public static void WriteAjaxResult(bool success, string result)
    {
        HttpContext.Current.Response.Expires = 0;
        HttpContext.Current.Response.CacheControl = "no-cache";
        HttpContext.Current.Response.AddHeader("Pragma", "no-cache");
        HttpContext.Current.Response.Write(result);
    }

    public static void SetManagerSessions()
    {
        HttpContext.Current.Session["memID"] = GetAppSetting("managerId");
        HttpContext.Current.Session["memName"] = GetAppSetting("managerName");
        HttpContext.Current.Session["memNickName"] = GetAppSetting("managerNickName");
    }

    public static string convertToJsonString(object data)
    {
        return JsonConvert.SerializeObject(data);
    }

    /// <summary>
    /// 取出KMAPI的設定
    /// </summary>
    /// <param name="key"></param>
    /// <returns></returns>
    public static string GetKMAPISetting(string key)
    {
        NameValueCollection nvc = (NameValueCollection)ConfigurationManager.GetSection("lamdbaapiSetting");
        return nvc[key];
    }

    /// <summary>
    /// 在頁面上呈現自訂MessageBox訊息
    /// </summary>
    /// <param name="objPage">Page</param>
    /// <param name="MsgText">messageg文字</param>
    /// <param name="Url">按下 alert 確定後，要導向的網址</param>
    public static void WindowAlertAndBack(Page objPage, string MsgText)
    {
        string altAlertMessage = "$AlertMessage$";
        string script = @"<script language=Javascript>alert(""$AlertMessage$"");history.go(-1);</script>";
        string finalScript = "";
        if (!String.IsNullOrEmpty(MsgText))
        {
            finalScript = script.Replace(altAlertMessage
                , HttpUtility.HtmlEncode(MsgText.Replace(Environment.NewLine, "\\n")));
            objPage.ClientScript.RegisterStartupScript(typeof(Page), "AlertMessage", finalScript);
        }
    }

    public static object DeserializeAjaxResult(string jsonString, Type type)
    {
        return JsonConvert.DeserializeObject(jsonString, type);
    }

    /// <summary>
    /// 建立文件類型清單
    /// </summary>
    /// <param name="documentClasses">文件類型List</param>
    /// <returns>Hashtable</returns>
    public static Hashtable GetBuildDocumentClassList(Page page)
    {
        Hashtable mappingTable = new Hashtable();
        if (CacheManager.Get("DocumentClassList") == null)
        {
            DocumentClassInfo[] documentClasses = GetAPIClient().documentClasses.GetAll(DocumentClass.Status.All);
            for (int i = 0; i < documentClasses.Length; i++)
            {
                DocumentClassInfo info = documentClasses[i];
                if (!mappingTable.ContainsKey(info.Uid))
                { mappingTable.Add(info.Uid.ToString(), info); }
            }
            CacheManager.Insert("DocumentClassList", mappingTable, new DateTime().AddMinutes(30));
        }

        return (Hashtable)CacheManager.Get("DocumentClassList");
    }

    /// <summary>
    /// 取得預設的APIClient
    /// </summary>
    /// <returns>RestRpcClient</returns>
    public static RestRpcClient GetAPIClient()
    {
        return GetAPIClient(GetKMAPISetting("APIUrl"),
               GetKMAPISetting("APIKey"),
               GetKMAPISetting("APIActor"),
               GetKMAPISetting("APIFormat"),
               int.MaxValue);
    }

    /// <summary>
    /// 取得自定義的APIClient
    /// </summary>
    /// <param name="url">APIUrl路徑</param>
    /// <param name="key">APIKey</param>
    /// <param name="actor">使用者</param>
    /// <param name="format">回傳格式</param>
    /// <param name="pageSize">筆數</param>
    /// <returns>RestRpcClient</returns>
    public static RestRpcClient GetAPIClient(string url, string key, string actor, string format, int pageSize)
    {
        RestRpcClient api = new RestRpcClient(new Uri(url));
        api.key = key;
        api.actor = actor;
        api.format = GSS.Vitals.API.Client.Serializer.ParseFormat(format);
        api.pageSize = pageSize;
        return api;
    }

    public static string GetAPIParameterString()
    {
        return "guid=" + Guid.NewGuid().ToString().ToLower()
            + "&format=" + GetKMAPISetting("APIFormat")
            + "&tid=0"
            + "&who=" + GetKMAPISetting("APIActor")
            + "&pi=" + 0
            + "&ps=" + 10
            + "&api_key=" + GetKMAPISetting("APIKey");

    }

    /// <summary>
    /// 建立TreeView所需的XML資料
    /// </summary>
    /// <param name="doc">XML文件</param>
    /// <param name="currentNodeID">目前點選的節點</param>
    /// <param name="collection">暫存XML容器</param>
    public static void GenerateXMLDataSource(XmlDocument doc, int currentNodeID, string rootDisplayName, string requestURL, ArrayList collection)
    {
        RestRpcClient api = GetAPIClient();
        IList tempCategoryList = new ArrayList();
        //#取出目前傳入節點的Children
        CategoryInfoPagedCollection baseCategoryInfoList = api.categories.GetChildren(
         currentNodeID, false);
        //#取出目前傳入節點的CategoryInfo
        CategoryInfo currentCategoryInfo = api.categories.Get(currentNodeID, false);

        //#判斷是否為RootNode
        bool isRootNode = (currentCategoryInfo.CategoryId.ToString() !=
            WebUtility.GetAppSetting("CategoryTreeRootNode") ? false : true);

        //#建立XML父節點
        XmlNode xmlNode = CreateCategoryInfoXmlNode(doc, currentCategoryInfo
            , (isRootNode ? rootDisplayName : null), requestURL);

        //#建立Children節點
        if (baseCategoryInfoList != null)
        {
            foreach (CategoryInfo categoryInfo in baseCategoryInfoList.Elements)
            {
                XmlNode xmlCategory = CreateCategoryInfoXmlNode(doc, categoryInfo, null, requestURL);
                if (collection.Count == 0)
                { xmlNode.AppendChild(xmlCategory); }
                else
                {
                    XmlNode currentNode = (XmlNode)collection[0];
                    if (categoryInfo.CategoryId.ToString() == currentNode.Attributes["ID"].Value)
                    { xmlNode.AppendChild(currentNode); }
                    else
                    { xmlNode.AppendChild(xmlCategory); }
                }
            }
        }

        if (collection.Count == 0)
        { collection.Add(xmlNode); }
        else
        { collection[0] = xmlNode; }

        //#判斷是否Root決定Recursive
        if (!isRootNode)
        {
            CategoryInfo parentCategoryInfo = api.categories.GetParent(currentNodeID, false);
            GenerateXMLDataSource(doc, parentCategoryInfo.CategoryId, rootDisplayName, requestURL, collection);
        }
        else
        { doc.AppendChild((XmlNode)collection[0]); }
    }

    /// <summary>
    /// 建立XML節點
    /// </summary>
    /// <param name="xmlDocument">XML文件</param>
    /// <param name="categoryInfo">類別資訊</param>
    /// <returns>XML節點</returns>
    private static XmlNode CreateCategoryInfoXmlNode(XmlDocument xmlDocument, CategoryInfo categoryInfo, string rootDisplayName, string requestURL)
    {
        XmlNode xmlNode = xmlDocument.CreateNode(XmlNodeType.Element, "node", "");
        XmlAttribute attributeName = xmlDocument.CreateAttribute("name");
        XmlAttribute attributeUrl = xmlDocument.CreateAttribute("Url");
        XmlAttribute attributeID = xmlDocument.CreateAttribute("ID");
        attributeName.Value = (rootDisplayName != null ? rootDisplayName : categoryInfo.DisplayName);
        attributeUrl.Value = string.Format(requestURL, categoryInfo.CategoryId, (WebUtility.GetStringParameter("ActorType") == "000" ? "002" : WebUtility.GetStringParameter("ActorType")));
        attributeID.Value = categoryInfo.CategoryId.ToString();
        xmlNode.Attributes.Append(attributeName);
        xmlNode.Attributes.Append(attributeUrl);
        xmlNode.Attributes.Append(attributeID);
        return xmlNode;
    }

    /// <summary>
    /// 將入口網000/001/002/003轉換成系統內的分眾參數
    /// </summary>
    /// <param name="actorType">actorType</param>
    /// <returns>string</returns>
    public static string ConvertGroupsReadingInInter(string actorType)
    {
        string MapperGroupsReadingInInter = string.Empty;
        switch (actorType)
        {
            case "000":
                MapperGroupsReadingInInter = "A,B";
                break;
            case "001":
                MapperGroupsReadingInInter = "A";
                break;
            case "002":
                MapperGroupsReadingInInter = "B";
                break;
            case "003":
                MapperGroupsReadingInInter = "C";
                break;
            default:
                break;
        }
        return MapperGroupsReadingInInter;
    }

    public static string ParseISO8601SimpleString(string inputString, DateTimeKind kind)
    {
        TimeZoneInfo defaultTimeZoneInfo = TimeZoneInfo.Local;
        string[] tz = (defaultTimeZoneInfo.BaseUtcOffset.ToString() + "@" + defaultTimeZoneInfo.Id.ToString()).Split('@');
        string h = tz[0].Split(':')[0];
        string m = tz[0].Split(':')[1];

        TimeSpan tp = new TimeSpan(Convert.ToInt32(h), Convert.ToInt32(m), 0);

        int yyyy = int.Parse(inputString.Substring(0, 4));

        int MM = int.Parse(inputString.Substring(4, 2));

        int dd = int.Parse(inputString.Substring(6, 2));

        int HH = int.Parse(inputString.Substring(8, 2));

        int mm = int.Parse(inputString.Substring(10, 2));

        int ss = int.Parse(inputString.Substring(12, 2));

        DateTime dt = new DateTime(yyyy, MM, dd, HH, mm, ss, kind);

        return dt.Add(tp).ToString("yyyy/MM/dd HH:mm:ss");
    }

    public static string GetThisPageName()
    {
        return (HttpContext.Current.Request.FilePath.Substring(HttpContext.Current.Request.FilePath.LastIndexOf('/') + 1));
    }

    public static bool checkParam(string s)
    {
        string[] stringArray = new string[] { "<", ">", "%3C", "%3E", ";", "%27", "'","\"","%22" , ";" , "%2d" };
        foreach (string temp in stringArray)
        {
            if (s.IndexOf(temp) != -1)
                return true;
        }
        return false;
    }
    public static bool checkParam(string[] s)
    {
        foreach (string temp in s)
        {
            if (checkParam(temp))
                return true;
        }
        return false;
    }
    public static bool checkInt(string s)
    {
        int temp;
        if (int.TryParse(s, out temp))
            return true;
        return false;
    }
    public static bool checkInt(string[] s)
    {
        foreach (string temp in s)
        {
            if (checkInt(temp))
                return true;
        }
        return false;
    }


    public static string GetMyAreaLinks(MyAreaLink link)
    {
        string str = @"<ul>
                          <li class='{0}'><a href=""/knowledge/myknowledge_record.aspx""><span>我的紀錄</span></a></li>
                          <li class='{1}'><a href=""/knowledge/myknowledge_question_lp.aspx""><span>我的發問</span></a></li>
                          <li class='{2}'><a href=""/knowledge/myknowledge_discuss_lp.aspx""><span>我的討論</span></a></li>
                          <li class='{3}'><a href=""/knowledge/myknowledge_trace_lp.aspx""><span>我的追蹤</span></a></li>
                          <li class='{4}'><a href=""/knowledge/myknowledge_pedia.aspx""><span>我的小百科</span></a></li>
				          <li class='{5}'><a href=""/knowledge/myknowledge_QuestionResponse.aspx""><span>我反應的問題</span></a></li>
                          <li class='{6}'><a href=""/recommand/recommand_Mylist.aspx""><span>我推薦的文章</span></a></li>
                        <!--
                          <li class='{7}'><a href=""/knowledge/myknowledge_tagsubject.aspx""><span>我的標籤</span></a></li>
                          <li class='{8}'><a href=""/knowledge/myknowledge_read.aspx""><span>我的閱讀記錄</span></a></li>
                          <li class='{9}'><a href=""/knowledge/myknowledge_keep.aspx""><span>我的收藏</span></a></li>
                        -->
                    </ul>";


        //設定選定的Tab
        string[] arr = new string[Enum.GetValues(typeof(MyAreaLink)).Length];
        arr[(int)link] = "current";        

        return string.Format(str, arr);
    }    

}


public enum MyAreaLink
{
    myknowledge_record = 0,
    myknowledge_question_lp = 1,
    myknowledge_discuss_lp = 2,
    myknowledge_trace_lp = 3,
    myknowledge_pedia = 4,
    myknowledge_QuestionResponse = 5,
    recommand_Mylist = 6,
    myknowledge_tags = 7,
    myknowledge_read = 8,
    myknowledge_keep = 9
}

