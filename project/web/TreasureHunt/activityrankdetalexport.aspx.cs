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

public partial class TreasureHunt_activityrankdetalexport : System.Web.UI.Page
{
    private TreasureHunt treasureHunt;
    protected void Page_Load(object sender, EventArgs e)
    {
        Response.Clear();
        Response.Buffer = true;
        Response.Charset = "utf-8";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + DateTime.Now.ToString("yyyymmddhhmmss") + "_treasureTop.xls");
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8");
        Response.ContentType = "application/ms-excel";
        SetExcelData();
    }

    private void SetExcelData()
    {
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>");
        treasureHunt = new TreasureHunt("");
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
        totalsuits = (WebUtility.GetStringParameter("totalsuits", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("totalsuits", string.Empty));
        int totalsuitUpper = 99;
        totalsuitUpper =  (WebUtility.GetStringParameter("totalsuitupper", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("totalsuitupper", string.Empty));
        
        int totalsetsLowerBound = (WebUtility.GetStringParameter("totalsetslowerbound", string.Empty) == "") ? 0 : Convert.ToInt32(WebUtility.GetStringParameter("totalsetslowerbound", string.Empty));
        int totalsetsUpperBound = (WebUtility.GetStringParameter("totalsetsupperbound", string.Empty) == "") ? 999 : Convert.ToInt32(WebUtility.GetStringParameter("totalsetsupperbound", string.Empty));
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
        bool reverse = false;
        string reverseString = WebUtility.GetStringParameter("reverse", string.Empty).ToLower();
        reverse = (reverseString.CompareTo("false") != 0 || reverseString.CompareTo("true") != 0) ? Convert.ToBoolean(reverseString) : false;
        //第一次活動所以給1
        treasureHunt.SetActivity(avtivityId);
        bool onlyGetBox = false;
        if (WebUtility.GetStringParameter("states", string.Empty).ToLower() == "getbox")
            onlyGetBox = true;
        Dictionary<int, TreasureHunt.Treasure> treasureDetal = treasureHunt.GetTreasureDetal();
        IList topList = treasureHunt.GetTopAndQuery(pageNumber, pageSize, avtivityId, queryMember, treasureId, totalsuits, totalsuitUpper, probability, totalsetsLowerBound, totalsetsUpperBound, startTime, endTime, reverse, onlyGetBox);

        if (topList.Count > 0)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<table border=\"1\">");
            sb.Append("<tr><td>&nbsp;</td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">帳號</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">姓名｜暱稱</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">已收集套數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">抽獎次數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">有效閱讀數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">尋寶獲得寶物數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">獲得種類&數量</font></td>");
            if (avtivityId == 3)
            {
                sb.Append("<td align=\"center\"><font face=\"新細明體\">交換次數</font></td>");
                sb.Append("<td align=\"center\"><font face=\"新細明體\">實際獲得寶物數</font></td>");
                sb.Append("<th scope=\"col\" width=\"15%\">實際擁有種類</th>");
            }
            sb.Append("</tr>");


            for (int i = 0; i < topList.Count; i++)
            {
                sb.Append("<tr><td align=\"center\">" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                TreasureHunt.Treasure_Top topObj = new TreasureHunt.Treasure_Top();
                topObj = (TreasureHunt.Treasure_Top)topList[i];
                sb.Append("<td align=\"center\">" + topObj.LoginId + "</td>");
                if (!string.IsNullOrEmpty(topObj.NickName) && !string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + topObj.RealName + "(" + topObj.NickName + ")" + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.NickName))
                {
                    sb.Append("<td align=\"center\">" + "(" + topObj.NickName + ")" + "</td>");
                }
                else if (!string.IsNullOrEmpty(topObj.RealName))
                {
                    sb.Append("<td align=\"center\">" + topObj.RealName + "</td>");
                }
                else
                {
                    sb.Append("<td>&nbsp;</td>");
                }
                sb.Append("<td align=\"center\">"+ topObj.TotalSuit + "</td>");

                sb.Append("<td align=\"center\">" + topObj.Lottery + "</td>");
                sb.Append("<td align=\"center\">" + topObj.TotalReadCount + "</td>");
                sb.Append("<td align=\"center\">" + topObj.Total_Treasure.ToString() + "</td>");
                string ss = "<td align=\"center\">";
                //取得各user的寶物細項
                Dictionary<int, int> treasureCount = treasureHunt.GetUsetAllKindTreasureCount(avtivityId, topObj.AccountId, startTime, endTime);
                foreach (KeyValuePair<int, TreasureHunt.Treasure> userItem in treasureDetal)
                {
                    TreasureHunt.Treasure treasure = userItem.Value;
                    if (treasure.TreasureName.Length > 3)
                    {
                        ss += treasure.TreasureName.Substring(3, 1) + ":";
                    }
                    else
                    {
                        ss += treasure.TreasureName + ":";
                    }
                    if (treasureCount.ContainsKey(userItem.Key))
                    {
                        ss += treasureCount[userItem.Key].ToString();
                    }
                    else
                    {
                        ss += "0";
                    }
                    ss += "&nbsp;";
                }
                ss += "</td>";
				sb.Append(ss);
                if (avtivityId == 3)
                {
                    sb.Append("<td align=\"center\">" + topObj.Change_Treasure.ToString() + "</td>");
                    sb.Append("<td align=\"center\">" + topObj.Real_Treasure.ToString() + "</td>");
                
                    Dictionary<int, int> usersTreasure = treasureHunt.GetUsersTreasure(topObj.AccountId);
                    ss = "<td align=\"center\">";
                    foreach (KeyValuePair<int, TreasureHunt.Treasure> userItem in treasureDetal)
                    {
                        TreasureHunt.Treasure treasure = userItem.Value;
                        if (treasure.TreasureName.Length > 3)
                        {
                            ss += treasure.TreasureName.Substring(3, 1) + ":";
                        }
                        else
                        {
                            ss += treasure.TreasureName + ":";
                        }
                        if (usersTreasure.ContainsKey(userItem.Key))
                        {
                            ss += usersTreasure[userItem.Key].ToString();
                        }
                        else
                        {
                            ss += "0";
                        }
                        ss += "&nbsp;";
                    }
                    ss += "</td>";
                    sb.Append(ss);
                }

                sb.Append("</tr>");

            }
            sb.Append("</table>");
            Response.Write(sb);
        }
        else
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("<table border=\"1\">");
            sb.Append("<tr><td>&nbsp;</td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">帳號</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">姓名｜暱稱</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">已收集套數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">抽獎次數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">有效閱讀數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">尋寶獲得寶物數</font></td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">獲得種類&數量</font></td>");
            if (avtivityId == 3)
            {
                sb.Append("<td align=\"center\"><font face=\"新細明體\">交換次數</font></td>");
                sb.Append("<td align=\"center\"><font face=\"新細明體\">實際獲得寶物數</font></td>");
            }
            sb.Append("</tr>");
            sb.Append("</table>");
            Response.Write(sb);

        }


    }
}
