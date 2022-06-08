using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;

public partial class Century_TabText : System.Web.UI.UserControl
{
    private string iCTUnitEvent, iCTUnitPic, iCTUnitStory;
    public DateTime? atDate { get { return atDate; } set { atDate = value; } }

    protected void Page_Load(object sender, EventArgs e)
    {
        iCTUnitEvent = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitEvent"].ToString();
        iCTUnitPic = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitPic"].ToString();
        iCTUnitStory = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnit"].ToString();
        // 設定特定日期;


        // 透過檔案取得檔名
        string strPhysicalPath = System.IO.Path.GetFileName(Request.PhysicalPath);
        string sTemplate = "{0}<a href={1}>{2}</a></li>";

        GroupText.Text = "<ul class='group'>";

        if (checkUse(iCTUnitEvent))
        {
            GroupText.Text += string.Format(sTemplate
                                , SetActivity(strPhysicalPath, "Event_List.aspx")
                                , "Event_List.aspx", "大事紀");
            if (addHistory())
            {
                GroupText.Text += string.Format(sTemplate
                                    , SetActivity(strPhysicalPath, "History_List.aspx")
                                    , "History_List.aspx", "歷史上的今天");
            }
        }
        if (checkUse(iCTUnitPic))
        {
            GroupText.Text += string.Format(sTemplate
                                , SetActivity(strPhysicalPath, "Picture_List.aspx", "Picture_Detail.aspx")
                                , "Picture_List.aspx", "珍貴老照片");
        }
        if (checkUse(iCTUnitStory))
        {
        GroupText.Text += string.Format(sTemplate
                            , SetActivity(strPhysicalPath, "Story_List.aspx", "Story_Detail.aspx")
                            , "Story_List.aspx", "農業故事");
        }
        GroupText.Text += "</ul>";

        switch (strPhysicalPath)
        {
            case "Event_List.aspx":
                labSubject.Text = "百年農業發展史上的：";
                break;
            case "History_List.aspx":
                labSubject.Text = "百年農業發展史上的今天：";
                break;
            default:
                break;
        }
    }

    // 判斷當日是否有歷史事件
    protected bool addHistory()
    {
        // today{month, day}
        int[] NowDate = { DateTime.Now.Month, DateTime.Now.Day };
        string strQueryScript = @"SELECT    HistoryList.*, CuDtGeneric.xBody
                                  FROM HistoryList INNER JOIN CuDtGeneric
                                        ON CuDtGeneric.icuitem=HistoryList.gicuitem
                                  WHERE CuDtGeneric.iCTUnit= @iCTUnit
                                        AND CuDtGeneric.fCTUPublic  = 'Y' AND HistoryList.Month = @Month
                                        AND HistoryList.Day = @Day
                                  ORDER BY HistoryList.[Year], HistoryList.[Month], HistoryList.[day]";
        using (var reader = SqlHelper.ReturnReader("ODBCDSN", strQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN", "@iCTUnit", "@iCTUnit", iCTUnitEvent),
             DbProviderFactories.CreateParameter("ODBCDSN", "@Month", "@Month", NowDate[0].ToString()),
             DbProviderFactories.CreateParameter("ODBCDSN", "@Day", "@Day", NowDate[1].ToString())))
        {
            if (reader.HasRows)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    // 判斷特定日期是否有歷史事件
    protected bool addHistory(int Month, int Day)
    {
        string strQueryScript = @"SELECT    HistoryList.*, CuDtGeneric.xBody
                                  FROM HistoryList INNER JOIN CuDtGeneric
                                        ON CuDtGeneric.icuitem=HistoryList.gicuitem
                                  WHERE CuDtGeneric.iCTUnit= @iCTUnit
                                        AND CuDtGeneric.fCTUPublic  = 'Y' AND HistoryList.Month = @Month
                                        AND HistoryList.Day = @Day
                                  ORDER BY HistoryList.[Year], HistoryList.[Month], HistoryList.[day]";
        using (var reader = SqlHelper.ReturnReader("ODBCDSN", strQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN", "@iCTUnit", "@iCTUnit", iCTUnitEvent),
             DbProviderFactories.CreateParameter("ODBCDSN", "@Month", "@Month", Month.ToString()),
             DbProviderFactories.CreateParameter("ODBCDSN", "@Day", "@Day", Day.ToString())))
        {
            if (reader.HasRows)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    // 設定使用中的Class
    private string SetActivity(string strPhysicalPath, params string[] href)
    {
        foreach (var item in href)
        {
            if (strPhysicalPath == item)
            {
                return "<li class='activity'>";
            }
        }
        return "<li>";
    }

    // 檢查主題單元是否啟用中
    private bool checkUse(string iCTUnitID)
    {
        string strQueryScript = @"SELECT inUse FROM CtUnit WHERE CtUnitID = @iCTUnitID";
        using (var reader = SqlHelper.ReturnReader("ODBCDSN", strQueryScript,
            DbProviderFactories.CreateParameter("ODBCDSN", "@iCTUnitID", "@iCTUnitID", iCTUnitID)))
        {
            if (reader.Read())
            {
                if (reader["inUse"].ToString().ToUpper() == "Y")
                {
                    return true;
                }
            }
        }
        return false;
    }

}