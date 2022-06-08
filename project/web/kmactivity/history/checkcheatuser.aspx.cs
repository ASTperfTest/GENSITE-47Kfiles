using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Data;

public partial class kmactivity_history_checkcheatuser : System.Web.UI.Page
{
    private HistoryPicture historyPicture;
    protected void Page_Load(object sender, EventArgs e)
    {
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        historyPicture = new HistoryPicture("");
        GetOverGetQuestion();
        GetOverGetScore();
        GetOverGetDailyScore();
    }
    private void GetOverGetQuestion()
    {
        string ss = "";
        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>日期</th><th>取題數</th></tr>";
        int orders = int.Parse(DropDownList1.SelectedValue);

        var dt = historyPicture.GetOverGetQuestion(orders);
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                ss += "<tr>";
                ss += "<td>";
                ss += dr["login_id"].ToString();
                ss += "</td>";
                ss += "<td>";
                ss += dr["realname"].ToString();
                ss += "</td>";
                ss += "<td>";
                ss += dr["createdate"].ToString();
                ss += "</td>";
                ss += "<td>";
                ss += dr["usercount"].ToString();
                ss += "</td>";
                ss += "</tr>";
            }
            ss += "</table>";
        }
        else
        {
            ss += "<tr><td colspan=\"4\">沒有單日答題超過30題的使用者</td></tr></table>";
        }
        TableText.Text = ss;
    }

    private void GetOverGetScore()
    {
        string ss = "";
        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>日期</th><th>當日總得分</th></tr>";
        int orders = int.Parse(DropDownList2.SelectedValue);
        var dt = historyPicture.GetOverGetScore(orders);
        if (dt.Rows.Count > 0)
        {
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    ss += "<tr>";
                    ss += "<td>";
                    ss += dr["login_id"].ToString();
                    ss += "</td>";
                    ss += "<td>";
                    ss += dr["realname"].ToString();
                    ss += "</td>";
                    ss += "<td>";
                    ss += dr["createdate"].ToString();
                    ss += "</td>";
                    ss += "<td>";
                    ss += dr["score"].ToString();
                    ss += "</td>";
                    ss += "</tr>";
                }
                ss += "</table>";
            }
        }
        else
        {
            ss += "<tr><td colspan=\"4\">沒有單日得分超過180分的使用者</td></tr></table>";
        }
        Label1.Text = ss;
    }


    private void GetOverGetDailyScore()
    {
        string ss = "";
        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>日期</th><th>當日總得分</th></tr>";
        string sql = @" select * from HISTORY_PICTURE_LOG ee 
                         left join account a on a.account_id = ee.account_id
                        where score >6  {0}
        ";
        int orders = int.Parse(DropDownList3.SelectedValue);
        string orderstring = "";
        switch (orders)
        {
            case 0:
                orderstring = " order by ee.create_date ,login_id ";
                break;
            case 1:
                orderstring = " order by login_id ,ee.create_date ";
                break;
            default:
                break;

        }
        sql = string.Format(sql, orderstring);
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString",sql);
        if (dt.Rows.Count > 0)
        {
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    ss += "<tr>";
                    ss += "<td>";
                    ss += dr["login_id"].ToString();
                    ss += "</td>";
                    ss += "<td>";
                    ss += dr["realname"].ToString();
                    ss += "</td>";
                    ss += "<td>";
                    ss += dr["createdate"].ToString();
                    ss += "</td>";
                    ss += "<td>";
                    ss += dr["score"].ToString();
                    ss += "</td>";
                    ss += "</tr>";
                }
                ss += "</table>";
            }
        }
        else
        {
            ss += "<tr><td colspan=\"4\">沒有單次得分超過6分的使用者</td></tr></table>";
        }
        Label2.Text = ss;
    }
}