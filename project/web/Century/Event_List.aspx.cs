using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using GSS.Vitals.COA.Data;
using System.Globalization;

public partial class Century_Events_Events_List : System.Web.UI.Page
{
    private int Month, iCTUnit;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["month"] == null || string.IsNullOrEmpty(Request.QueryString["month"].ToString()))
        {
            Month = DateTime.Now.Month;
        }
        else
        {
            Month = Convert.ToInt32(Request.QueryString["month"].ToString());
        }
        iCTUnit = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitEvent"].ToString());
        myDBinit();
    }

    protected void myDBinit()
    {
        // 先查詢日期並過濾重覆。
        string strQueryMonth = @"SELECT DISTINCT Month, Day FROM HISTORYLIST WHERE Month = @Month ORDER BY [day]";
        // 再依日期取得歷年事件。
        string strQueryString = @"SELECT    HistoryList.*, CuDtGeneric.xBody
                                  FROM HistoryList INNER JOIN CuDtGeneric
                                        ON CuDtGeneric.icuitem=HistoryList.gicuitem
                                  WHERE CuDtGeneric.iCTUnit= @iCTUnit
                                        AND CuDtGeneric.fCTUPublic  = 'Y' AND HistoryList.Month = @Month
                                        AND HistoryList.Day = @Day
                                  ORDER BY HistoryList.[Year], HistoryList.[Month], HistoryList.[day]";

        using (var Datereader = SqlHelper.ReturnReader("ODBCDSN", strQueryMonth,
            DbProviderFactories.CreateParameter("ODBCDSN", "@Month", "@Month", Month)))
        {
            labList.Text = "<table class='ListTable' width='100%'>";
            if (Datereader.HasRows)
            {
                // 日期的Loop
                labList.Text += "<tr>";
                while (Datereader.Read())
                {
                    // <div class='date_box'>
                    // 事件無日期不顯示日曆外框.. Modified By Leo 2011-10-31
                    if (string.IsNullOrEmpty(Datereader["Day"].ToString()))
                    {
                        labList.Text += "<td valign='top' width='5%'>&nbsp;</td>";
                    }
                    else
                    {
                        labList.Text += "<td valign='top' width='5%'><div class='date_box'><div class='date_box_month'>" + ConvertMonth(Month)
                        + "</div><div class='date_box_day'>" + Datereader["Day"].ToString() + "</div></td>";
                    }

                    // 當日所發生的歷史事件(無日期)
                    if (string.IsNullOrEmpty(Datereader["Day"].ToString()))
                    {
                        string strNullString = @"SELECT    HistoryList.*, CuDtGeneric.xBody
                                  FROM HistoryList INNER JOIN CuDtGeneric
                                        ON CuDtGeneric.icuitem=HistoryList.gicuitem
                                  WHERE CuDtGeneric.iCTUnit= @iCTUnit
                                        AND CuDtGeneric.fCTUPublic  = 'Y' AND HistoryList.Month = @Month
                                        AND HistoryList.Day IS NULL
                                  ORDER BY HistoryList.[Year], HistoryList.[Month], HistoryList.[day]";
                        using (var Contentreader = SqlHelper.ReturnReader("ODBCDSN", strNullString,
                        DbProviderFactories.CreateParameter("ODBCDSN", "@iCTUnit", "@iCTUnit", iCTUnit),
                        DbProviderFactories.CreateParameter("ODBCDSN", "@Month", "@Month", Month)))
                        {
                            if (Contentreader.HasRows)
                            {
                                string temp = "";   // 暫存上一筆的年份
                                labList.Text += "<td>";
                                while (Contentreader.Read())
                                {
                                    if (temp != Contentreader["Year"].ToString())
                                    {
                                        labList.Text += "<font color='#cabF00'>" + Contentreader["Year"].ToString() + "年:</font><br/>"
                                            + "<ul><a name='" + Contentreader["gicuitem"].ToString() + "'></a><li>" 
                                            + Contentreader["xBody"] + "</li></ul>";
                                    }
                                    else
                                    {
                                        labList.Text += "<ul><a name='" + Contentreader["gicuitem"].ToString() + "'></a><li>" 
                                            + Contentreader["xBody"] + "</li></ul>";
                                    }
                                    temp = Contentreader["Year"].ToString();
                                }
                                labList.Text += "</td>";
                            }
                        }
                    }
                    // 當日所發生的歷史事件(有日期)
                    else
                    {
                        using (var Contentreader = SqlHelper.ReturnReader("ODBCDSN", strQueryString,
                            DbProviderFactories.CreateParameter("ODBCDSN", "@iCTUnit", "@iCTUnit", iCTUnit),
                            DbProviderFactories.CreateParameter("ODBCDSN", "@Month", "@Month", Month),
                            DbProviderFactories.CreateParameter("ODBCDSN", "@Day", "@Day", Datereader["Day"].ToString())))
                        {
                            if (Contentreader.HasRows)
                            {
                                string temp = "";   // 暫存上一筆的年份
                                labList.Text += "<td>";
                                while (Contentreader.Read())
                                {
                                    if (temp != Contentreader["Year"].ToString())
                                    {
                                        labList.Text += "<font color='#cabF00'>" + Contentreader["Year"].ToString() + "年:</font><br/>"
                                            + "<ul><a name='" + Contentreader["gicuitem"].ToString() + "'></a><li>"
                                            + Contentreader["xBody"] + "</li></ul>";
                                    }
                                    else
                                    {
                                        labList.Text += "<ul><a name='" + Contentreader["gicuitem"].ToString() + "'></a><li>" 
                                            + Contentreader["xBody"] + "</li></ul>";
                                    }
                                    temp = Contentreader["Year"].ToString();
                                }
                                labList.Text += "</td>";
                            }
                        }
                    }
                    labList.Text += "</tr>";
                }
            }
            labList.Text += "</table>";
        }
    }

    // 月份輸除成英文縮寫
    protected string ConvertMonth(int Month)
    {
        DateTime Result = Convert.ToDateTime("1900/" + Month + "/1");
        DateTimeFormatInfo myDTFI = new CultureInfo("en-US", false).DateTimeFormat;
        return Result.ToString("MMM", myDTFI);
    }
}