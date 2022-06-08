using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Collections;

public partial class kmactivity_history_historylogexport : System.Web.UI.Page
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
        string accountId = WebUtility.GetStringParameter("accountid", string.Empty).ToLower();
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>");
        int pageSize = 99999;
        int pageNumber = 1;
        string sql = @" select * from account where account_id = @account_id
        ";
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@account_id", "@account_id", accountId));
        string loginId = dt.Rows[0]["login_id"].ToString();
        IList objlist = historyPicture.GerUserQuestioninfo(loginId, pageSize, pageNumber);
        string ss = "<table><tr><th></th><th>日期</th><th>答案</th><th>答題狀況</th><th>得分</th></tr>";
        foreach (HistoryPicture.HistoryQuestionInfo obj in objlist)
        {
            ss += "<tr>";
            ss += "<td>" + obj.RowNumber + "</td>";
            ss += "<td>" + obj.CreateDate + "</td>";
            ss += "<td>" + obj.Stitle + "</td>";
            if (obj.State == 3)
            {
                ss += "<td>答題成功</td>";
            }
            else
            {
                ss += "<td>答題失敗</td>";
            }
            string imageUrl = "<img src=\"/public/History/" + historyPicture.GetCurrentPic(obj.Icuitem) + "\" width=\"30px\" height=\"30px\" /> ";

            ss += "<td>" + obj.Score + "</td>";
            ss += "</tr>";
        }
        ss += "</table>";
        Response.Write(ss);
    }
}