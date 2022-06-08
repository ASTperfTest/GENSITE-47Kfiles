using System;
using System.Text;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Web.Configuration;

public class Pager
{
    public int PageIndex { get; private set; }
    public int PageSize { get; private set; }
    public int TotalCount { get; private set; }
    public int TotalPages { get; private set; }
    public StringBuilder PageOptions { get; private set; }
    public StringBuilder PageSizeOptions { get; private set; }
    private string[] PageSizeList = { "10", "30", "50" };

    public Pager(int pageIndex, int pageSize, int totalCount, string[] pageSizeList)
    {
        PageIndex = pageIndex;
        PageSize = pageSize == 0 ? int.MaxValue : pageSize;
        if (pageSizeList != null)
            PageSizeList = pageSizeList;

        TotalCount = totalCount;
        TotalPages = (int)Math.Ceiling(TotalCount / (double)PageSize);

        PageOptions = new StringBuilder("");
        for (int i = 0; i < TotalPages; i++)
            PageOptions.Append("<option value=\"" + (i + 1).ToString() + "\"" + (i == PageIndex ? " selected=\"selected\"" : "") + ">" + (i + 1).ToString() + "</option>");

        PageSizeOptions = new StringBuilder("");
        foreach (var x in PageSizeList)
            PageSizeOptions.Append("<option value=\"" + x + "\"" + (x == pageSize.ToString() ? " selected=\"selected\"" : "") + ">" + x + "</option>");
    
    }

    public bool HasPreviousPage
    {
        get
        {
            return (PageIndex > 0);
        }
    }

    public bool HasNextPage
    {
        get
        {
            return (PageIndex + 1 < TotalPages);
        }
    }
}

public static class myFunc
{
    public static string GetWebSetting(string xName, string xType)
    {
        Configuration rootWebConfig = WebConfigurationManager.OpenWebConfiguration("/");
        switch (xType)
        {
            case "appSettings":
                if (rootWebConfig.AppSettings.Settings[xName] != null)
                    return rootWebConfig.AppSettings.Settings[xName].Value;
                break;
            case "connectionStrings":
                if (rootWebConfig.ConnectionStrings.ConnectionStrings[xName] != null)
                    return rootWebConfig.ConnectionStrings.ConnectionStrings[xName].ConnectionString;
                break;
        }
        return null;
    }
    public static void GenErrMsg(String errMessage, String errAction)
    {
        if ((errMessage.Length > 0) || (errAction.Length > 0))
        {
            HttpContext.Current.Response.Write("<script language=\"JavaScript\" type=\"text/javascript\">");
            if (errMessage.Length > 0)
            {
                HttpContext.Current.Response.Write("alert('" + errMessage + "');");
            }
            if (errAction.Length > 0)
            {
                HttpContext.Current.Response.Write(errAction);
            }
            HttpContext.Current.Response.Write("</script>");
            HttpContext.Current.Response.End();
        }
    }
    public static string GetDefaultImg(string x, string types, string type)
    {
        const string imgURL = "/public/data/"; 
        string result = x;
        if (string.IsNullOrEmpty(x))
        {
            if (types == "0")
            {
                switch (type)
                {
                    case "0":
                        result = "image/果.png";
                        break;
                    case "1":
                        result = "image/菜.png";
                        break;
                    case "2":
                        result = "image/花.png";
                        break;
                    case "3":
                        result = "image/穀.png";
                        break;
                }
            }
            else if (types == "1")
            {
                result = "image/漁.png";
            }
        }
        else
        {
            result = string.Concat(imgURL, x);
        }
        return result;
    }
}