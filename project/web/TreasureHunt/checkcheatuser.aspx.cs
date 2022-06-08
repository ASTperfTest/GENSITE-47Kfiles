using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class TreasureHunt_checkcheatuser : System.Web.UI.Page
{
    protected int activityid = 3;
    private TreasureHunt treasureHunt;
    protected void Page_Load(object sender, EventArgs e)
    {
        treasureHunt = new TreasureHunt("");
        treasureHunt.SetActivity(activityid);
        SetCheckUserPackage();
        SetdailyGet();
        SetdailyChange();
    }
    private void SetCheckUserPackage()
    {
        Dictionary<int, TreasureHunt.Treasure> treasureDetal = treasureHunt.GetTreasureDetal();
        DataTable dt = treasureHunt.CheckUserPackage(activityid);
        string ss = "";
        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>寶物名稱</th><th>擁有寶物數</th><th>log紀錄裡的寶物數</th></tr>";
        if (dt.Rows.Count > 0)
        {
            foreach(DataRow dr in dt.Rows)
            {
                ss+="<tr>";
                ss += "<td align=\"center\">" + dr["login_id"].ToString() + "</td>";
                if (!string.IsNullOrEmpty(dr["nickname"].ToString()) && !string.IsNullOrEmpty(dr["realname"].ToString()))
                {
                    ss += "<td align=\"center\">" + dr["realname"].ToString() + "(" + dr["nickname"].ToString() + ")" + "</td>";
                }
                else if (!string.IsNullOrEmpty(dr["nickname"].ToString()))
                {
                    ss += "<td align=\"center\">" + "(" + dr["nickname"].ToString() + ")" + "</td>";
                }

                ss += "<td align=\"center\">" + treasureDetal[int.Parse(dr["treasure_id"].ToString())].TreasureName + "</td>";
                ss += "<td align=\"center\">" + dr["piece"].ToString() + "</td>";
                ss += "<td align=\"center\">" + dr["nowcount"].ToString() + "</td>";
                ss+="</tr>";
            }
        }
        else
        {
            ss += "<tr><td colspan=\"5\">查無資料</td></tr>";
        }
        ss += "</tr></table>";
        TableText.Text = ss;
    }
    private void SetdailyGet()
    {
        DataTable dt = treasureHunt.CheckDailyGetSet(activityid);
        string ss = "";
        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>換出or換得</th><th>取得數量</th><th>日期</th></tr>";
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                ss += "<tr>";
                ss += "<td align=\"center\">" + dr["login_id"].ToString() + "</td>";
                if (!string.IsNullOrEmpty(dr["nickname"].ToString()) && !string.IsNullOrEmpty(dr["realname"].ToString()))
                {
                    ss += "<td align=\"center\">" + dr["realname"].ToString() + "(" + dr["nickname"].ToString() + ")" + "</td>";
                }
                else if (!string.IsNullOrEmpty(dr["nickname"].ToString()))
                {
                    ss += "<td align=\"center\">" + "(" + dr["nickname"].ToString() + ")" + "</td>";
                }
                ss += "<td align=\"center\">" + dr["treasure_id"].ToString() + "</td>";
                ss += "<td align=\"center\">" + dr["daily_count"].ToString() + "</td>";
                ss += "<td align=\"center\">" + dr["modify_date"].ToString() + "</td>";
                ss += "<td align=\"center\">" + dr["nickname"].ToString() + "</td>";
                ss += "</tr>";
            }
        }
        else
        {
            ss += "<tr><td colspan=\"6\">查無資料</td></tr>";
        }
        ss += "</tr></table>";
        Label1.Text = ss;
    }

    private void SetdailyChange()
    {
        DataTable dt = treasureHunt.CheckDailyChangeSet(activityid);
        string ss = "";
        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>寶物id</th><th>數量</th><th>日期</th></tr>";
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                ss += "<tr>";
                ss += "<td align=\"center\">" + dr["login_id"].ToString() + "</td>";
                if (!string.IsNullOrEmpty(dr["nickname"].ToString()) && !string.IsNullOrEmpty(dr["realname"].ToString()))
                {
                    ss += "<td align=\"center\">" + dr["realname"].ToString() + "(" + dr["nickname"].ToString() + ")" + "</td>";
                }
                else if (!string.IsNullOrEmpty(dr["nickname"].ToString()))
                {
                    ss += "<td align=\"center\">" + "(" + dr["nickname"].ToString() + ")" + "</td>";
                }
                if (dr["states"].ToString() == "3")
                {
                    ss += "<td align=\"center\">換出</td>";
                }
                else if (dr["states"].ToString() == "4")
                {
                    ss += "<td align=\"center\">換得</td>";
                }
                ss += "<td align=\"center\">" + dr["daily_count"].ToString() + "</td>";
                ss += "<td align=\"center\">" + dr["create_date"].ToString() + "</td>";
                ss += "</tr>";
            }
        }
        else
        {
            ss += "<tr><td colspan=\"6\">查無資料</td></tr>";
        }
        ss += "</tr></table>";
        Label2.Text = ss;
    }
}