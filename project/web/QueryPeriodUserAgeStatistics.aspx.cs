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

public partial class QueryPeriodUserAgeStatistics : System.Web.UI.Page
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
                        int maxAge = 100;

                        StatisticsTitle.Visible = true;
                        DataTable resultTable = GetStatisticsResult2(inputBeginDate, inputEndDate, maxAge);

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

    private static DataTable GetStatisticsResult2(string inputBeginDate, string inputEndDate, int maxAge)
    {
        string sql = @"Select 
                        DATEDIFF(YEAR,CONVERT(VARCHAR,Mem.birthday,111),GETDATE()) / 5 as [年齡],
                        count(Distinct MemberId) as [人數], sum(loginInterCount) as [人次]
                        FROM MemberGradeLogin AS GRADE LEFT JOIN Member AS Mem 
                        ON GRADE.memberId=Mem.account 
                        WHERE loginDate >= '#StartDate#' AND loginDate < '#EndDate#' 
                        AND ISNULL(Mem.birthday,'') > '#MinBirthday#'
						AND ISNULL(Mem.birthday,'') < '#MaxBirthday#'
                        AND LEN(LTRIM(Mem.birthday))=8
                        AND ISDATE(Mem.birthday)=1
                        Group By DATEDIFF(YEAR,CONVERT(VARCHAR,Mem.birthday,111),GETDATE()) / 5
                        Order By [年齡]      ";

        sql = sql.Replace("#StartDate#", Convert.ToDateTime(inputBeginDate).ToString("yyyy-MM-dd"));
        sql = sql.Replace("#EndDate#", Convert.ToDateTime(inputEndDate).AddDays(1).ToString("yyyy-MM-dd"));
        sql = sql.Replace("#MinBirthday#", System.DateTime.Now.AddYears(- maxAge).ToString("yyyy-MM-dd"));
		sql = sql.Replace("#MaxBirthday#", System.DateTime.Now.ToString("yyyy-MM-dd"));

        string webConfigConnectionString =
                System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

        DataTable TempTable = new DataTable();
        SqlDataAdapter DA = new SqlDataAdapter(sql, webConfigConnectionString);
        DA.Fill(TempTable);
       
        DataTable resultTable = new DataTable();
        resultTable.Columns.Add("年齡");
        resultTable.Columns.Add("人數");
        resultTable.Columns.Add("人次");

        int P = 0;
        string[] rowDataArr = new string[] { "", "", "" };
        for (int i = 0; i * 5 < maxAge; i += 1)
        {
            DataRow dr = resultTable.NewRow();

            if (TempTable.Rows.Count > P && Convert.ToInt32(TempTable.Rows[P]["年齡"]) == i)
            {
                rowDataArr = new string[] { (i * 5) + " ~ " + ((i + 1) * 5 - 1), TempTable.Rows[P]["人數"].ToString(), TempTable.Rows[P]["人次"].ToString() };
                P += 1;
            }
            else
            {
                rowDataArr = new string[] { (i * 5) + " ~ " + ((i + 1) * 5 - 1), "0", "0" };
            }

            dr.ItemArray = rowDataArr;

            resultTable.Rows.Add(dr);
        }

        return resultTable;
    }

	//取得統計結果
    private static DataTable GetStatisticsResult(string inputBeginDate, string inputEndDate, int maxAge)
    {
        string sql = @"WITH tt AS 
                        (
                            SELECT COUNT(1) as LoginMemberCount
                            FROM MemberGradeLogin AS GRADE
                            LEFT JOIN Member AS Mem 
                            ON GRADE.memberId=Mem.account
                            WHERE CONVERT(varchar, loginDate, 111) <= @endDate --迄日
                            AND CONVERT(varchar, loginDate, 111) >= @beginDate --起日
                            AND ISNULL(Mem.birthday,'')>'18000000'
                            AND LEN(LTRIM(Mem.birthday))=8 
                            AND ISDATE(Mem.birthday)=1
                            AND DATEDIFF(YEAR,CONVERT(VARCHAR,Mem.birthday,111),GETDATE())>=@rangeStart
                            AND DATEDIFF(YEAR,CONVERT(VARCHAR,Mem.birthday,111),GETDATE())<@rangeEnd
                            GROUP BY account--以人為一單位
                        )
                        SELECT count(1) FROM tt";

        DataTable resultTable = new DataTable();
        resultTable.Columns.Add("年齡");
        resultTable.Columns.Add("人數");
        resultTable.Columns.Add("人次");


        using (SqlConnection sqlConn = new SqlConnection())
        {
            string webConfigConnectionString =
                System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

            sqlConn.ConnectionString = webConfigConnectionString;
            sqlConn.Open();

			int rangeEnd = 0;
			int memberCount = 0;
			string[] rowDataArr = new string[] {"",""};
            for (int rangeStart = 0; rangeStart < maxAge; rangeStart += 5)
            {
                rangeEnd = rangeStart + 5;

                SqlCommand cmd = new SqlCommand(sql, sqlConn);
                cmd.Parameters.AddWithValue("@beginDate", inputBeginDate);
                cmd.Parameters.AddWithValue("@endDate", inputEndDate);
                cmd.Parameters.AddWithValue("@rangeStart", rangeStart);
                cmd.Parameters.AddWithValue("@rangeEnd", rangeEnd);

                memberCount = (int)cmd.ExecuteScalar();

                DataRow dr = resultTable.NewRow();

                rowDataArr = new string[] { rangeStart + " ~ " + rangeEnd, memberCount.ToString(), "0" };
                dr.ItemArray = rowDataArr;
                resultTable.Rows.Add(dr);

            }
        }
        return resultTable;
    }

  


    

}
