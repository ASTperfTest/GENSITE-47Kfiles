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

public partial class TreasureHunt_lotterydetailexport : System.Web.UI.Page
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
        TreasureHunt treasureHunt = new TreasureHunt("");
        int pageSize = 99999;
        int pageNumber = 1;
        int avtivityId = 1;
        avtivityId = (WebUtility.GetStringParameter("avtivityid", string.Empty) == "") ? 1 : Convert.ToInt32(WebUtility.GetStringParameter("avtivityid", string.Empty));
        string giftId = "";
        giftId = (WebUtility.GetStringParameter("giftid", string.Empty) == "") ? "" : WebUtility.GetStringParameter("giftid", string.Empty);
        string queryMember = "";
        queryMember = (WebUtility.GetStringParameter("querymember", string.Empty) == "") ? "" : WebUtility.GetStringParameter("querymember", string.Empty);
        queryMember = HttpUtility.UrlDecode(queryMember);
        IList lotteryTopList = treasureHunt.GetGiftVoteUsers(avtivityId, giftId, queryMember, pageNumber, pageSize);
        int total = treasureHunt.GetVoteCountById(avtivityId, giftId);
       

        if (lotteryTopList.Count > 0)
        {
            string link = "";
            StringBuilder sb = new StringBuilder();

            sb.Append("<table border=\"1\">");
            sb.Append("<tr><td>&nbsp;</td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">帳號</td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">姓名｜暱稱</td>");
            sb.Append("<td align=\"center\"><font face=\"新細明體\">換取次數</td>");
            sb.Append("</tr>");


            for (int i = 0; i < lotteryTopList.Count; i++)
            {
                sb.Append("<tr><td align=\"center\">" + ((pageSize * (pageNumber - 1)) + i + 1).ToString() + "</td>");
                TreasureHunt.Treasure_Top topObj = new TreasureHunt.Treasure_Top();
                topObj = (TreasureHunt.Treasure_Top)lotteryTopList[i];
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
                sb.Append("<td align=\"center\">"+ topObj.Lottery + "</td>");

                sb.Append("</tr>");

            }
            sb.Append("</table>");
            Response.Write(sb.ToString());
        }
        else
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<table class=\"type02\" summary=\"結果條列式\">");
            sb.Append("<tr><th scope=\"col\" width=\"5%\">&nbsp;</th>");
            sb.Append("<th scope=\"col\" width=\"10%\">帳號</th>");
            sb.Append("<th scope=\"col\" width=\"15%\">姓名｜暱稱</th>");
            sb.Append("<th scope=\"col\" width=\"10%\">換取次數</th>");
            sb.Append("</tr>");
            sb.Append("</table>");
            Response.Write(sb.ToString());
        }

    }
}
