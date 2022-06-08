using System;

/// <summary>
/// Json 的日期格式與.Net DateTime類型的轉換工具。
/// "/Date(1242357713797+0800)/"
/// </summary>
public class JsonDateTimeUtility
{
    /// <summary>
    /// 轉換為.Net DateTime。
    /// </summary>
    /// <param name="jsonDate">Json 的日期格式字串</param>
    /// <returns>轉換後的日期</returns>
    public static DateTime ToDateTime( string jsonDate )
    {
        string value = jsonDate.Substring( 6, jsonDate.Length - 8 );
        DateTimeKind kind = DateTimeKind.Utc;
        int index = value.IndexOf( '+', 1 );
        if( index == -1 )
            index = value.IndexOf( '-', 1 );
        if( index != -1 )
        {
            kind = DateTimeKind.Local;
            value = value.Substring( 0, index );
        }
        long javaScriptTicks = long.Parse( value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture );
        long InitialJavaScriptDateTicks = ( new DateTime( 1970, 1, 1, 0, 0, 0, DateTimeKind.Utc ) ).Ticks;
        DateTime utcDateTime = new DateTime( ( javaScriptTicks * 10000 ) + InitialJavaScriptDateTicks, DateTimeKind.Utc );
        DateTime dateTime;
        switch( kind )
        {
            case DateTimeKind.Unspecified:
                dateTime = DateTime.SpecifyKind( utcDateTime.ToLocalTime(), DateTimeKind.Unspecified );
                break;
            case DateTimeKind.Local:
                dateTime = utcDateTime.ToLocalTime();
                break;
            default:
                dateTime = utcDateTime;
                break;
        }
        return dateTime;
    }
    /// <summary>
    /// 轉換為Json 的日期格式字串。
    /// </summary>
    /// <param name="dt">.Net DateTime</param>
    /// <returns>轉換後的字串</returns>
    public static string ToJson( DateTime dt )
    {
        string dtString = "";
        //日期及時間
        long InitialJavaScriptDateTicks = ( new DateTime( 1970, 1, 1, 0, 0, 0, DateTimeKind.Utc ) ).Ticks;
        long ActualTicks = dt.Ticks;
        long DeltaTicks = ( ActualTicks - InitialJavaScriptDateTicks ) / 10000;

        //位移部份
        string op = "";
        string hr = "00";
        string mn = "00";
        TimeSpan offset = DateTimeOffset.Now.Offset;
        hr = offset.TotalHours.ToString( "00" );
        mn = ( offset.TotalMinutes - ( offset.TotalHours * 60 ) ).ToString( "00" );
        if( offset >= TimeSpan.Zero )
        {
            op = "+";
        }
        else
        {
            op = "-";
        }

        switch( dt.Kind )
        {
            case DateTimeKind.Utc:
                dtString = "\\/Date(" + DeltaTicks.ToString() + ")\\/";
                break;
            case DateTimeKind.Local:
                dtString = "\\/Date(" + DeltaTicks.ToString() + op + hr + mn + ")\\/";
                break;
            case DateTimeKind.Unspecified:
                dtString = "\\/Date(" + DeltaTicks.ToString() + op + hr + mn + ")\\/";
                break;
        }
        return dtString;
    }
}
