using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;
using System.Data;

public partial class CheckCheatUser : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        GetOverGetQuestion(); // 超過5題者
        GetOverGetScore(); // 超過10分者
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
    private void GetOverGetQuestion()
    {
        string ss = "";
        bool overFive = false;
        int orders = int.Parse(DropDownList1.SelectedValue);
        string sql = @"select B.*, ACCOUNT.REALNAME, ACCOUNT.NICKNAME 
                       from (  select *, count(gametime) as Counting 
                               from (  select LOGIN_ID, convert(nvarchar,gametime,111) as gametime
                                       from GAMEHistory
                                       where picState='Y' or picState='F'
                               ) as A 
                               group by A.LOGIN_ID,A.gametime 
                       ) as B
                       left join ACCOUNT on ACCOUNT.LOGIN_ID=B.LOGIN_ID ";
        if (DropDownList1.SelectedValue == "0")
            sql += "order by B.gametime desc";
        else if (DropDownList1.SelectedValue == "1")
            sql += "order by B.LOGIN_ID";


        DataTable dt = SqlHelper.GetDataTable("PuzzleConnString", sql);
        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>日期</th><th>答題數</th></tr>";

        foreach (DataRow dr in dt.Rows)
        { 
            if ( int.Parse(dr["Counting"].ToString()) > 5 ) 
            {
                overFive = true;
                ss += "<tr>";
                ss += "<td>";
                ss += dr["LOGIN_ID"].ToString();
                ss += "</td>";
                ss += "<td>";
                ss += DealName( dr["REALNAME"], dr["NICKNAME"] );
                ss += "</td>";
                ss += "<td>";
                ss += dr["gametime"].ToString();
                ss += "</td>";
                ss += "<td>";
                ss += dr["Counting"].ToString();
                ss += "</td>";
                ss += "</tr>";
            }
        }

        if (!overFive)
        {
            ss += "<tr><td colspan=\"4\">沒有單日完成超過5題的使用者</td></tr></table>";
        }

        ss += "</table>";
        TableText.Text = ss;
    }

    private void GetOverGetScore()
    {
        string ss = "";
        int orders = int.Parse(DropDownList2.SelectedValue);
        string sql = @"select * from
                       (
                         select LOGIN_ID, REALNAME, NICKNAME, gamedate,sum(S) as score
                         from 
                         (
                           select A.*, (case difficult when 'H' then Counting*2 when 'E' then Counting else 0 end) as S
                           from (
                                  select ACCOUNT.LOGIN_ID, REALNAME, NICKNAME, convert(nvarchar,gametime,111) as gamedate, count(picState) as Counting, difficult
                                  from ACCOUNT 
                                  left join GAMEHistory on ACCOUNT.LOGIN_ID = GAMEHistory.LOGIN_ID 
                                  where picState='Y'
                                  group by ACCOUNT.LOGIN_ID, REALNAME, NICKNAME, convert(nvarchar,gametime,111), difficult
                           ) A 
                         ) B
                         group by gamedate,LOGIN_ID,NICKNAME,REALNAME
                       ) C
                       where score > 10";
        
        if (orders == 0)
            sql = sql + "order by gamedate desc";
        else if (orders == 1)
            sql = sql + "order by LOGIN_ID";

        ss += "<table width=\"500px\" class=\"type02\">";
        ss += "<tr><th>帳號</th><th>姓名/暱稱</th><th>日期</th><th>當日總得分</th></tr>";

        DataTable dt = SqlHelper.GetDataTable("PuzzleConnString", sql);
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
                    ss += DealName(dr["REALNAME"], dr["NICKNAME"]);
                    ss += "</td>";
                    ss += "<td>";
                    ss += dr["gamedate"].ToString();
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
            ss += "<tr><td colspan=\"4\">沒有單日得分超過10分的使用者</td></tr></table>";
        }
        Label1.Text = ss;
    }

}