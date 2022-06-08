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
using GSS.Vitals.COA.Data;

public partial class GameRankExport : System.Web.UI.Page
{
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

        string queryMember = "";
        queryMember = (WebUtility.GetStringParameter("querymember", string.Empty) == "") ? "" : WebUtility.GetStringParameter("querymember", string.Empty);
        queryMember = HttpUtility.UrlDecode(queryMember);
        int treasureId = 0;
        treasureId = (WebUtility.GetStringParameter("treasureid", string.Empty) == "") ? 0 : Convert.ToInt32(WebUtility.GetStringParameter("treasureid", string.Empty));

        int totalsuits = 0;
        totalsuits = (WebUtility.GetStringParameter("answercountlow", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("answercountlow", string.Empty));
        int totalsuitUpper = 990;
        totalsuitUpper = (WebUtility.GetStringParameter("answercountupp", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("answercountupp", string.Empty));

        int totalsetsLowerBound = (WebUtility.GetStringParameter("scoreUpper", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("scoreUpper", string.Empty));
        int totalsetsUpperBound = (WebUtility.GetStringParameter("scoreLowerBound", string.Empty) == "") ? -1 : Convert.ToInt32(WebUtility.GetStringParameter("scoreLowerBound", string.Empty));
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

        string sql = @"
        select  ROW_NUMBER() OVER(order by TotalGrade desc,B.Counting)as Row,GH.LOGIN_ID,
        isnull(B.Counting,0) as Counting,isnull(B.TotalGrade,0) as TotalGrade,
		isnull(C.useenergy,0) as useenergy,  mm.realname,isnull(mm.nickname,'') as nickname,mm.email,ac.Energy,ac.GetEnergy,
        case(isnull(mm.status,'Y') ) when 'Y' then 0 else 1 end as status from (
        select login_id from  GAMEHistory 
        where ( picState='F' or picState='Y' ) {0}
        group by login_id 
        ) GH
        left join 
        (
	        select  LOGIN_ID,sum(Counting) as Counting,sum (score) as  TotalGrade
                from (
                select LOGIN_ID,count(picState) as Counting,difficult,case difficult when 'E' then 1 * count(picState) else 2 * count(picState) end as score

                 from GAMEHistory where picState='Y' {0}
		        
                group by LOGIN_ID,difficult
                ) A 
                group by LOGIN_ID 
        ) B on B.LOGIN_ID = GH.LOGIN_ID 
        left  join 
                (
	                 select LOGIN_ID,sum(useenergy)-count(pic_id) as useenergy from (
        select LOGIN_ID,pic_id,count(picState) as useenergy

                 from GAMEHistory where  ( picState='N' or picState='Y' ) {0}
                group by LOGIN_ID,pic_id
        ) G 
        group by LOGIN_ID
                ) C on C.login_id = GH.login_id 
left  join ACCOUNT ac on ac.login_id = GH.login_id
left  join mGIPcoanew..member mm on mm.account = GH.login_id
where 1=1
        ";
        if (queryMember != "")
        {
            sql += " and  (mm.account like @login_id or mm.REALNAME like @login_id or mm.NICKNAME like @login_id) ";
        }
        if (totalsetsLowerBound != -1 && totalsetsUpperBound != -1)
        {
            sql += " and TotalGrade between @scoreLowerBound and @scoreUpper ";
        }
        if (totalsuits != -1 && totalsuitUpper != -1)
        {
            sql += " and Counting between @answercountlow and @answercountupp  ";
        }
        string timestring = "";
        if (startTime != "" && endTime != "")
        {
            if (DateTime.TryParseExact(startTime, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp) && DateTime.TryParseExact(endTime, "yyyy/MM/dd", null, System.Globalization.DateTimeStyles.None, out timeTemp))
            {
                timestring = " and gametime between @startTime and dateadd(\"d\",1,@endTime)  ";
            }
        }
        sql = string.Format(sql, timestring);
        DataTable dt = SqlHelper.GetDataTable("PuzzleConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", "%" + queryMember + "%"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scoreLowerBound", "@scoreLowerBound", totalsetsUpperBound),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scoreUpper", "@scoreUpper", totalsetsLowerBound),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountlow", "@answercountlow", totalsuits),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountupp", "@answercountupp", totalsuitUpper),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@startTime", "@startTime", startTime),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@endTime", "@endTime", endTime));
        StringBuilder sb = new StringBuilder();
        sb.Append("<table class=\"type02\" width=\"60%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" >");
        sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
        sb.Append("<th scope=\"col\" width=\"5%\">氨舦</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">眀腹</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">﹎际嘿</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">笆眔だ</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">ЧΘ计</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">ㄏノ翴计</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">ゼㄏノ翴计</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">羆翴计</th></tr>");
        if (dt.Rows.Count > 0)
        {
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sb.Append("<tr><td align=\"center\">" + dt.Rows[i][0] + "</td>");
                if (dt.Rows[i][10] != null && dt.Rows[i][10].ToString() == "1")
                {
                    sb.Append("<td align=\"center\">氨舦</td>"); // 氨舦
                }
                else
                {
                    sb.Append("<td align=\"center\"></td>"); // 氨舦
                }
                sb.Append("<td align=\"center\">" + dt.Rows[i][1] + "</td>"); // 眀腹
                sb.Append("<td align=\"center\">" + DealName(dt.Rows[i][6], dt.Rows[i][5]) + "</td>"); // ﹎际嘿
                sb.Append("<td align=\"center\">" + dt.Rows[i][3] + "</td>"); // 眔だ
                sb.Append("<td align=\"center\">" + dt.Rows[i][2] + "</td>"); // 氮肈计
                sb.Append("<td align=\"center\">" + dt.Rows[i][4] + "</td>"); // ㄏノ翴计
                sb.Append("<td align=\"center\">" + dt.Rows[i][8] + "</td>"); // ゼㄏノ翴计
                sb.Append("<td align=\"center\">" + dt.Rows[i][9] + "</td></tr>"); // 羆翴计
            }
        }
        sb.Append("</table>");
        Response.Write(sb.ToString());
    }

    protected string DealName(object realname, object nickname)
    {
        string str = "";
        if (nickname != null && !string.IsNullOrEmpty(nickname.ToString()))
        {
            return nickname.ToString();
        }
        else
        {
            str = realname.ToString();
            return str;
        }

    }
}