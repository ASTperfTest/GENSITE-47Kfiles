using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.SqlClient;

public partial class QueryUserLoginDetail : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        
    }
	//按下「開始查詢」按鈕
    protected void Query_Click(object sender, EventArgs e)
    {
        string inputBeginDate =TextBoxBeginDate.Text.Trim();
        string inputEndDate = TextBoxEndDate.Text.Trim();
        
        if (inputBeginDate != "" && inputEndDate != "")
        {
            try
            {
                DateTime beginDate = Convert.ToDateTime(inputBeginDate);
                DateTime endDate = Convert.ToDateTime(inputEndDate);

                if (endDate > beginDate)
                {
                    TimeSpan timeInterval = (endDate - beginDate);

                    if (timeInterval.TotalDays < 366)
                    {
                        StatisticsTitle.Visible = true;
                        
                        DataTable resultTable = GetStatisticsResult2(inputBeginDate, inputEndDate);

                        GridViewOfQueryResult.DataSource = resultTable;
                        GridViewOfQueryResult.DataBind();
                    }
                    else
                    {
                        Response.Write("<script>alert('開始與結束日期區間不可超過一年，以免影響效能。' )</script>");
                        StatisticsTitle.Visible = false;
                        GridViewOfQueryResult.DataBind();
                    }
                }
                else
                {
                    Response.Write("<script>alert('開始日期需小於結束日期。' )</script>");
                    StatisticsTitle.Visible = false;
                    GridViewOfQueryResult.DataBind();
                }
            }
            catch (Exception ex)
            {
                Response.Write(ex.ToString());
                Response.Write("<script>alert('日期格式不正確。' )</script>");
                StatisticsTitle.Visible = false;
                GridViewOfQueryResult.DataBind();
            }    
        }
        else
        {
            Response.Write("<script>alert('日期不可空白。' )</script>");
            StatisticsTitle.Visible = false;
            GridViewOfQueryResult.DataBind();
        }
                     
    }

    private DataTable GetStatisticsResult2(string inputBeginDate, string inputEndDate)
    {
        string sql = @"
            SELECT      memberId
                        , max(member.realname) as realname
                        , max(member.nickname) as nickname
                        , sum(LoginInterCount) as loginCount 
                        , max(LoginInterDate) as LastLoginDate
            FROM        MemberGradeLogin
            inner Join	member on MemberGradeLogin.MemberId = Member.Account
            WHERE     logindate between '#StartDate#' and cast('#EndDate#' as datetime) + 1 and not LoginInterDate is null
            group by memberId
            ORDER BY max(LoginInterDate) DESC
        ";

        sql = sql.Replace("#StartDate#", Convert.ToDateTime(inputBeginDate).ToString("yyyy/MM/dd"));
        sql = sql.Replace("#EndDate#", Convert.ToDateTime(inputEndDate).AddDays(1).ToString("yyyy/MM/dd"));

        //Response.Write(sql);
        string webConfigConnectionString =
                System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

        DataTable TempTable = new DataTable();
        SqlDataAdapter DA = new SqlDataAdapter(sql, webConfigConnectionString);
        DA.Fill(TempTable);

        return TempTable;
    }

    

}
