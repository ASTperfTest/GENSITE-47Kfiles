using System;
using System.Data;
using System.Configuration;
using System.Text;
using System.Web;
using System.Web.Caching;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Security.Cryptography;

/// <summary>
/// CacheManager 的摘要描述
/// </summary>
public class CacheManager
{
    public CacheManager()
    { }

    public static string Insert(string cacheKeyPrefix, object value)
    {
        return Insert(cacheKeyPrefix, value, TimeSpan.FromSeconds(1800));
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="cacheKeyPrefix"></param>
    /// <param name="value"></param>
    /// <param name="absoluteExpiration"></param>
    /// <returns></returns>
    /// <remarks>Insert(cacheKeyPrefix, value, DateTime.UtcNow.AddMinutes(30))</remarks>
    public static string Insert(string cacheKeyPrefix, object value, DateTime absoluteExpiration)
    {
        return Insert(cacheKeyPrefix, value
            , absoluteExpiration, System.Web.Caching.Cache.NoSlidingExpiration
            , null);
    }

    public static string Insert(string cacheKeyPrefix, object value, TimeSpan slidingExpiration)
    {
        return Insert(cacheKeyPrefix, value
            , System.Web.Caching.Cache.NoAbsoluteExpiration, slidingExpiration
            , null);
    }

    public static string Insert(string cacheKeyPrefix, object value
        , DateTime absoluteExpiration, TimeSpan slidingExpiration
        , CacheItemRemovedCallback onRemoveCallback)
    {
        string cacheKey = cacheKeyPrefix;

        HttpContext.Current.Cache.Insert(cacheKey, value, null, absoluteExpiration,
            slidingExpiration, System.Web.Caching.CacheItemPriority.Default
            , onRemoveCallback);

        return cacheKey;
    }

    public static object Get(string cacheKey)
    {
        return HttpContext.Current.Cache.Get(cacheKey);
    }
}
