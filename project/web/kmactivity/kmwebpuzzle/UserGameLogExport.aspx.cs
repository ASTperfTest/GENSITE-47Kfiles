using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Text;
using System.Data;

public partial class kmactivity_kmwebpuzzle_UserGameLogExport : System.Web.UI.Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
        Response.Clear();
        Response.Buffer = true;
        Response.Charset = "utf-8";
        Response.AddHeader("Content-Disposition", "attachment;filename=" + DateTime.Now.ToString("yyyymmddhhmmss") + "_userlog.xls");
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8");
        Response.ContentType = "application/ms-excel";
        ExportData();
    }

    private void ExportData()
    {
        string queryMember = "";
        queryMember = (WebUtility.GetStringParameter("querymember", string.Empty) == "") ? "" : WebUtility.GetStringParameter("querymember", string.Empty);
        queryMember = HttpUtility.UrlDecode(queryMember);
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>");
        string sql = @"
            select ROW_NUMBER() OVER(order by gametime DESC) as Row , login_id,
             picstate,GAMEHistory.pic_id,convert(nvarchar,gametime,111) as gametime,difficult,pic.pic_no,pic.pic_name
				,useenergy
              from GAMEHistory 
			left join PICDATA pic on pic.ser_no = pic_id
			left join (
				select pic_id,sum(useenergy)-count(pic_id) as useenergy from (
        select LOGIN_ID,pic_id,count(picState) as useenergy
                 from GAMEHistory where (picState='N' or picState='Y' )
			and login_id = @login_id
                group by pic_id,LOGIN_ID
        ) G 
        group by pic_id,LOGIN_ID
			) sg on sg.pic_id = GAMEHistory.pic_id
			where login_id = @login_id
            and (picstate = 'Y' or picstate = 'F')
            order by gametime desc
        ";
        var dt = SqlHelper.GetDataTable("PuzzleConnString", sql,
           DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", queryMember));
        StringBuilder sb = new StringBuilder();
        sb.Append("<table class=\"type02\" width=\"60%\" border=\"0\" cellpadding=\"0\" cellspacing=\"0\" >");
        sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
        sb.Append("<th scope=\"col\" width=\"5%\">日期</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">完成/放棄 </th>");
        sb.Append("<th scope=\"col\" width=\"15%\">難度</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">活動得分</th>");
        sb.Append("<th scope=\"col\" width=\"15%\">使用點數</th></tr>");
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                sb.Append("<tr><td align=\"center\">" + dr["Row"] + "</td>");
                sb.Append("<td align=\"center\">" + dr["gametime"] + "</td>");
                sb.Append("<td align=\"center\">" + GetStates(dr["picstate"]) + "</td>");
                sb.Append("<td align=\"center\">" + GetDifficult(dr["difficult"]) + "</td>");
                sb.Append("<td align=\"center\">" + GetScore(dr["difficult"], dr["picstate"]) + "</td>");
                sb.Append("<td align=\"center\">" + dr["useenergy"] + "</td></tr>");
            }
        }
        sb.Append("</table>");
        Response.Write(sb.ToString());

    }

    protected string GetStates(object difficult)
    {
        if (difficult.ToString() == "Y")
        {
            return "完成";
        }
        else
        {
            return "放棄";
        }
    }

    protected string GetDifficult(object difficult)
    {
        if (difficult.ToString() == "H")
        {
            return "完整";
        }
        else
        {
            return "簡單";
        }
    }


    protected string GetScore(object difficult, object picstates)
    {
        if (picstates.ToString() == "Y")
        {
            if (difficult.ToString() == "E")
            {
                return "1";
            }
            else
            {
                return "2";
            }
        }
        else
        {
            return "0";
        }
    }
}