using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Text;
using System.Collections.Generic;

public partial class kmactivity_history_activityrankdetalexport : System.Web.UI.Page
{
    private HistoryPicture historyPicture;
    protected void Page_Load(object sender, EventArgs e)
    {
        Response.Clear();
        Response.Buffer = true;
        Response.Charset = "utf-8";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + DateTime.Now.ToString("yyyymmddhhmmss") + "_treasureTop.xls");
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8");
        Response.ContentType = "application/ms-excel";
        historyPicture = new HistoryPicture("");
        SetExcelData();
    }
    private void SetExcelData()
    {
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>");
        int pageSize = 99999;
        int pageNumber = 1;
        double probability = -1;
        probability = (WebUtility.GetStringParameter("probability", string.Empty) == "") ? -1 : Convert.ToDouble(WebUtility.GetStringParameter("probability", string.Empty));

        int avtivityId = 1;
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 1 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", string.Empty));

        string queryMember = "";
        queryMember = (WebUtility.GetStringParameter("querymember", string.Empty) == "") ? "" : WebUtility.GetStringParameter("querymember", string.Empty);
        queryMember = HttpUtility.UrlDecode(queryMember);
        int treasureId = 0;
        treasureId = (WebUtility.GetStringParameter("treasureid", string.Empty) == "") ? 0 : Convert.ToInt32(WebUtility.GetStringParameter("treasureid", string.Empty));

        int totalsuits = 0;
        totalsuits = (WebUtility.GetStringParameter("answercountlow", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("answercountlow", string.Empty));
        int totalsuitUpper = 99;
        totalsuitUpper = (WebUtility.GetStringParameter("answercountupp", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("answercountupp", string.Empty));

        int totalsetsLowerBound = (WebUtility.GetStringParameter("scoreUpper", string.Empty) == "") ? 0 : Convert.ToInt32(WebUtility.GetStringParameter("scoreUpper", string.Empty));
        int totalsetsUpperBound = (WebUtility.GetStringParameter("scoreLowerBound", string.Empty) == "") ? 999 : Convert.ToInt32(WebUtility.GetStringParameter("scoreLowerBound", string.Empty));
        if (totalsetsLowerBound > totalsetsUpperBound)
        {
            int tempUpperBound = totalsetsUpperBound;
            totalsetsLowerBound = totalsetsUpperBound;
            totalsetsUpperBound = tempUpperBound;
        }

        string startTime = (WebUtility.GetStringParameter("starttime", string.Empty) == "") ? "" : WebUtility.GetStringParameter("starttime", string.Empty);
        string endTime = (WebUtility.GetStringParameter("endtime", string.Empty) == "") ? "" : WebUtility.GetStringParameter("endtime", string.Empty);
        DateTime timeTemp = new DateTime();
        if (!DateTime.TryParseExact(startTime, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp) || !DateTime.TryParseExact(endTime, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp))
        {
            startTime = "";
            endTime = "";
        }
        IList historyTop = historyPicture.GetTopUnsys(queryMember, "", totalsetsLowerBound, totalsetsUpperBound, totalsuits, totalsuitUpper, startTime, endTime, pageSize, pageNumber);
        StringBuilder sb = new StringBuilder();
        sb.Append("<table class=\"type02\" width=\"60%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" >");
        sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
        sb.Append("<th scope=\"col\" width=\"5%\">停權</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">帳號</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">活動得分</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">答題數</th></tr>");
        if (historyTop.Count > 0)
        {
            for (int i = 0; i < historyTop.Count; i++)
            {
                sb.Append("<tr><td>" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                HistoryPicture.HistoryTop topObj = (HistoryPicture.HistoryTop)historyTop[i];
                if (topObj.disable == true)
                {
                    sb.Append("<td align=\"center\">停權</td>");
                }
                else
                {
                    sb.Append("<td align=\"center\"></td>");
                }
                sb.Append("<td align=\"center\">"+ topObj.LoginId + "</td>");
                if (!string.IsNullOrEmpty(topObj.NickName))
                {
                    sb.Append("<td align=\"center\">" + topObj.NickName + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + (topObj.RealName.Substring(0, 1) + "＊" + topObj.RealName.Substring(topObj.RealName.Trim().Length - 1, 1)) + "</td>");
                }
                else
                {
                    sb.Append("<td>&nbsp;</td>");
                }
                sb.Append("<td align=\"center\">" + topObj.Score + "</td>");
                sb.Append("<td align=\"center\">" + topObj.Count + "</td>");
            }
        }
        sb.Append("</table>");
        Response.Write(sb.ToString());
    }
}