using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Globalization;

public partial class Century_History_List : System.Web.UI.Page
{
    private string iCTUnit, IsBirth;
    private int[] NowDate = { DateTime.Now.Month, DateTime.Now.Day };
    protected string strError = string.Empty;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["IsBirth"] != null && Session["memID"] !=null )
        {
            if (Session["IsBirth"] == "Y")
            {
                BirthdayLogin(Session["memID"].ToString());
            }
        }
        iCTUnit = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitEvent"].ToString();
        if (Request.QueryString["month"] == null && Request.QueryString["day"] == null)
        {
            myDBinit(NowDate[0].ToString(),NowDate[1].ToString());
        }
        else
        {
            myDBinit(Request.QueryString["month"], Request.QueryString["day"]);
        }
    }

//    protected void myDBinit()
//    {
//        // 再依日期取得歷年事件。
//        string strQueryString = @"SELECT    HistoryList.*, CuDtGeneric.xBody
//                                  FROM HistoryList INNER JOIN CuDtGeneric
//                                        ON CuDtGeneric.icuitem=HistoryList.gicuitem
//                                  WHERE CuDtGeneric.iCTUnit= @iCTUnit
//                                        AND CuDtGeneric.fCTUPublic  = 'Y' AND HistoryList.Month = @Month
//                                        AND HistoryList.Day = @Day
//                                  ORDER BY HistoryList.[Year], HistoryList.[Month], HistoryList.[day]";

//        // 當日所發生的歷史事件
//        using (var Contentreader = SqlHelper.ReturnReader("ODBCDSN", strQueryString,
//            DbProviderFactories.CreateParameter("ODBCDSN", "@iCTUnit", "@iCTUnit", iCTUnit),
//            DbProviderFactories.CreateParameter("ODBCDSN", "@Month", "@Month", NowDate[0].ToString()),
//            DbProviderFactories.CreateParameter("ODBCDSN", "@Day", "@Day", NowDate[1].ToString())))
//        {
//            if (Contentreader.HasRows)
//            {
//                string temp = "";   // 暫存上一筆的年份
//                labList.Text = @"<table class='ListTable'><tr><td valign='top'><div class='date_box'>
//                             <div class='date_box_month'>" + ConvertMonth(NowDate[0])
//                    + "</div><div class='date_box_day'>" + NowDate[1].ToString() + "</div></td><td>";
//                while (Contentreader.Read())
//                {
//                    if (temp != Contentreader["Year"].ToString())
//                    {
//                        labList.Text += "<font color='#cabF00'>" + Contentreader["Year"].ToString() + "年:</font><br/>"
//                            + "<ul><li>" + Contentreader["xBody"] + "</li></ul>";
//                    }
//                    else
//                    {
//                        labList.Text += "<ul><li>" + Contentreader["xBody"] + "</li></ul>";
//                    }
//                    temp = Contentreader["Year"].ToString();
//                }
//                labList.Text += "</td></tr></table>";
//            }
//            // 若無資料則導回大事紀
//            else
//            {
//                Response.Write(@"<script>alert('歷年的當日並無事紀，導回大事紀首頁');
//                                location.href('Event_List.aspx')</script>");
//                //Response.Redirect("Event_List.aspx");
//            }
//        }
//    }

    protected void myDBinit(string atMonth, string atDay)
    {
        // 再依日期取得歷年事件。
        string strQueryString = @"SELECT    HistoryList.*, CuDtGeneric.xBody
                                  FROM HistoryList INNER JOIN CuDtGeneric
                                        ON CuDtGeneric.icuitem=HistoryList.gicuitem
                                  WHERE CuDtGeneric.iCTUnit= @iCTUnit
                                        AND CuDtGeneric.fCTUPublic  = 'Y' AND HistoryList.Month = @Month
                                        AND HistoryList.Day = @Day
                                  ORDER BY HistoryList.[Year], HistoryList.[Month], HistoryList.[day]";

        // 當日所發生的歷史事件
        using (var Contentreader = SqlHelper.ReturnReader("ODBCDSN", strQueryString,
            DbProviderFactories.CreateParameter("ODBCDSN", "@iCTUnit", "@iCTUnit", iCTUnit),
            DbProviderFactories.CreateParameter("ODBCDSN", "@Month", "@Month", atMonth),
            DbProviderFactories.CreateParameter("ODBCDSN", "@Day", "@Day", atDay)))
        {
            if (Contentreader.HasRows)
            {
                string temp = "";   // 暫存上一筆的年份
                labList.Text = @"<table class='ListTable' width='100%'><tr><td valign='top' width='5%'><div class='date_box'>
                             <div class='date_box_month'>" + ConvertMonth(Convert.ToInt32(atMonth))
                    + "</div><div class='date_box_day'>" + atDay + "</div></td><td>";
                while (Contentreader.Read())
                {
                    if (temp != Contentreader["Year"].ToString())
                    {
                        labList.Text += "<font color='#cabF00'>" + Contentreader["Year"].ToString() + "年:</font><br/>"
                            + "<ul><li>" + Contentreader["xBody"] + "</li></ul>";
                    }
                    else
                    {
                        labList.Text += "<ul><li>" + Contentreader["xBody"] + "</li></ul>";
                    }
                    temp = Contentreader["Year"].ToString();
                }
                labList.Text += "</td></tr></table>";
            }
            // 若無資料則導回大事紀
            else
            {
//                Response.Write(@"<script>alert('歷年的當日並無事紀，導回大事紀首頁');
//                                location.href('Event_List.aspx')</script>");
                strError = @"<script>alert('歷年的當日並無事紀，導回大事紀首頁');
                                location.href='Event_List.aspx'</script>";
                //Response.Redirect("Event_List.aspx");
            }
        }
    }

    protected string ConvertMonth(int Month)
    {
        DateTime Result = Convert.ToDateTime("1900/" + Month + "/1");
        DateTimeFormatInfo myDTFI = new CultureInfo("en-US", false).DateTimeFormat;
        return Result.ToString("MMM", myDTFI);
    }

    protected void BirthdayLogin(string UserAccount)
    {
        string strQueryScript = "SELECT loginBirthdayDate FROM Member WHERE account = @account";
        using (var reader = SqlHelper.ReturnReader("ODBCDSN", strQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN", "@account", "@account", UserAccount)))
        {
            if (reader.HasRows)
            {
                if (reader.Read())
                {
                    if (!string.IsNullOrEmpty(reader["loginBirthdayDate"].ToString()))
                    {
                        // 前次登錄時間小於365天，直接中斷程式不執行 SQLScript。(程式最後只寫入SQL Script)
                        TimeSpan ts = DateTime.Now - Convert.ToDateTime(reader["loginbirthdayDate"].ToString());
                        if (Convert.ToInt32(ts.Days.ToString()) < 365)
                        { return; }
                    }
                    #region "取得生日卡KPI"
                    // 寫入登入日期(loginBirthdayDate)
                    string strMemberUpdateScript = "UPDATE Member SET loginBirthdayDate = @loginBirthdayDate WHERE account = @account";
                    SqlHelper.ExecuteNonQuery("ODBCDSN", strMemberUpdateScript,
                        DbProviderFactories.CreateParameter("ODBCDSN", "@loginBirthdayDate", "@loginBirthdayDate", DateTime.Now.ToShortDateString()),
                        DbProviderFactories.CreateParameter("ODBCDSN", "@account", "@account", UserAccount));

                    // 取得最新登入的一筆，Update [loginBirthdayGrade] to 1
                    string strMemLoginUpdate = @"UPDATE  MemberGradeLogin SET loginBirthdayGrade = 1 
                                     WHERE loginDate = (SELECT max(loginDate) FROM MemberGradeLogin
                                                        WHERE memberId = @memberId) AND 
                                           memberId = @memberId";
                    SqlHelper.ExecuteNonQuery("ODBCDSN", strMemLoginUpdate,
                        DbProviderFactories.CreateParameter("ODBCDSN", "@memberId", "@memberId", UserAccount));
                    string strCardKPIScript = "INSERT INTO BirthdayCardKPIRecord (account, getDate) VALUES (@account, @getDate)";
                    SqlHelper.ExecuteNonQuery("ODBCDSN",strCardKPIScript,
                        DbProviderFactories.CreateParameter("ODBCDSN", "@account", "@account", UserAccount),
                        DbProviderFactories.CreateParameter("ODBCDSN", "@getDate", "@getDate", DateTime.Now.ToShortDateString()));
                    #endregion
                }
            }
        }
    }
}