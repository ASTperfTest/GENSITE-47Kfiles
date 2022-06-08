using System;
using System.Collections.Generic;
using System.Linq;
using GenericRepository;
using System.Data.Linq;
using System.Web;
using System.Configuration;
using System.Web.Configuration;
using System.Text;

public interface IRepository : IGenericRepository { }
public class Repository : LSGenericRepository, IRepository
{
    public Repository(DataContext context)
        : base(context) { }
}

public class PaginatedList<T> : List<T>
{
    public int PageIndex { get; private set; }
    public int PageSize { get; private set; }
    public int TotalCount { get; private set; }
    public int TotalPages { get; private set; }
    public StringBuilder PageOptions { get; private set; }
    public StringBuilder PageSizeOptions { get; private set; }
    private string[] PageSizeList = { "0", "10", "30", "50" };

    public PaginatedList(IQueryable<T> source, int pageIndex, int pageSize, string[] pageSizeList)
    {
        PageIndex = pageIndex;
        PageSize = pageSize == 0 ? int.MaxValue : pageSize;
        if (pageSizeList != null)
            PageSizeList = pageSizeList;

        TotalCount = source.Count();
        TotalPages = (int)Math.Ceiling(TotalCount / (double)PageSize);

        this.AddRange(source.Skip(PageIndex * PageSize).Take(PageSize));

        PageOptions = new StringBuilder("");
        for (int i = 0; i < TotalPages; i++)
            PageOptions.Append("<option value=\"" + (i + 1).ToString() + "\"" + (i == PageIndex ? " selected=\"selected\"" : "") + ">" + (i + 1).ToString() + "</option>");

        PageSizeOptions = new StringBuilder("");
        foreach (var x in PageSizeList)
            PageSizeOptions.Append("<option value=\"" + x + "\"" + (x == pageSize.ToString() ? " selected=\"selected\"" : "") + ">" + (x == "0" ? "全部" : x) + "</option>");
    
    }

    public PaginatedList(IQueryable<T> source, int pageIndex, int pageSize, string[] pageSizeList, int dataCount)
    {
        PageIndex = pageIndex;
        PageSize = pageSize == 0 ? int.MaxValue : pageSize;
        if (pageSizeList != null)
            PageSizeList = pageSizeList;

        TotalCount = dataCount;
        TotalPages = (int)Math.Ceiling(TotalCount / (double)PageSize);

        this.AddRange(source.Skip(0 * PageSize).Take(PageSize));

        PageOptions = new StringBuilder("");
        for (int i = 0; i < TotalPages; i++)
            PageOptions.Append("<option value=\"" + (i + 1).ToString() + "\"" + (i == PageIndex ? " selected=\"selected\"" : "") + ">" + (i + 1).ToString() + "</option>");

        PageSizeOptions = new StringBuilder("");
        foreach (var x in PageSizeList)
            PageSizeOptions.Append("<option value=\"" + x + "\"" + (x == pageSize.ToString() ? " selected=\"selected\"" : "") + ">" + (x == "0" ? "全部" : x) + "</option>");

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

public static class StrFunc
{
    public static string Qty(string Str)
    {
        return (Str.Replace("'", "''"));
    }

    public static string NZQty(string Str)
    {
        return (Str == null ? "Null" : "'" + Qty(Str) + "'");
    }

    public static string uNZQty(string Str)
    {
        return (Str == null ? "Null" : "N'" + Qty(Str) + "'");
    }

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
        return "";
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

    // IsNumeric Function
    public static bool IsNumeric(object Expression)
    {
        // Variable to collect the Return value of the TryParse method.
        bool isNum;

        // Define variable to collect out parameter of the TryParse method. If the conversion fails, the out parameter is zero.
        double retNum;

        // The TryParse method converts a string in a specified style and culture-specific format to its double-precision floating point number equivalent.
        // The TryParse method does not generate an exception if the conversion fails. If the conversion passes, True is returned. If it does not, False is returned.
        isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
        return isNum;
    }
}