using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Data;
using System.Text.RegularExpressions;
public partial class QuerySearchKeywordStatistics : System.Web.UI.Page
{
    string webConfigConnectionString = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
    protected void Page_Load(object sender, EventArgs e)
    {


    }

    public static bool IsWholeNumber(String InputString)
    {
        return (InputString != string.Empty && !Regex.IsMatch(InputString, "[^0-9]"))
             ? true : false;
    }
    protected void btnQuery_Click(object sender, EventArgs e)
    {
        bool isDateOk = false;
        bool isNumberOk = false;
        #region 檢查
        if (tboxTagNumStart.Text.Trim() != "" && !IsWholeNumber(tboxTagNumStart.Text.Trim()))
        {
            Response.Write("<script>alert('關鍵字數量下限請輸入數字' )</script>");
            return;
        }
        if (tboxTagNumEnd.Text.Trim() != "" && !IsWholeNumber(tboxTagNumEnd.Text))
        {
            Response.Write("<script>alert('關鍵字數量上限請輸入數字' )</script>");
            return;
        }
        if (tboxDateStart.Text.Trim() != "")
        {
            DateTime dtTemp = DateTime.MinValue;
            if (!DateTime.TryParse(tboxDateStart.Text, out dtTemp) || tboxDateStart.Text.Trim().Length != 10)
            {
                Response.Write("<script>alert('您輸入的是不合法的日期,YYYY/MM/DD' )</script>");
                return;
            }
        }
        if (tboxDateEnd.Text.Trim() != "")
        {
            DateTime dtTemp = DateTime.MinValue;
            if (!DateTime.TryParse(tboxDateEnd.Text, out dtTemp) || tboxDateEnd.Text.Trim().Length != 10)
            {
                Response.Write("<script>alert('您輸入的是不合法的日期,YYYY/MM/DD' )</script>");
                return;
            }
        }
        if (tboxDateStart.Text != "" && tboxDateEnd.Text != "")
        {
            isDateOk = true;
            if (DateTime.Parse(tboxDateStart.Text) > DateTime.Parse(tboxDateEnd.Text))
            {
                Response.Write("<script>alert('開始日期不可大於結束日期' )</script>");
                return;
            }
        }

        if (tboxTagNumStart.Text != "" && tboxTagNumEnd.Text != "")
        {
            isNumberOk = true;
            if (int.Parse(tboxTagNumStart.Text.Trim()) > int.Parse(tboxTagNumEnd.Text.Trim()))
            {
                Response.Write("<script>alert('關鍵字下限不可大於上限' )</script>");
                return;
            }
        }

        if ((tboxTagNumStart.Text.Trim() != "" && tboxTagNumEnd.Text.Trim() == "") ||
            (tboxTagNumStart.Text.Trim() == "" && tboxTagNumEnd.Text.Trim() != ""))
        {
            Response.Write("<script>alert('關鍵字數量上限下限都要填寫' )</script>");
            return;
        }

        if ((tboxDateStart.Text.Trim() != "" && tboxDateEnd.Text.Trim() == "") ||
            (tboxDateStart.Text.Trim() == "" && tboxDateEnd.Text.Trim() != ""))
        {
            Response.Write("<script>alert('日期的開始與結束都要填寫' )</script>");
            return;
        }
        #endregion


        string SQL = "";
        SQL += "SELECT DISPLAY_NAME AS 關鍵字, USED_COUNT AS 點閱次數,LAST_USED AS 最後查詢日期  FROM TAG WHERE 1=1 ";
        //闗鍵字
        if (tboxTag.Text.Trim() != "")
        {
            SQL += "AND  DISPLAY_NAME LIKE '%" + tboxTag.Text.Trim() + "%'";
        }
        //日期限制(2者都選了後)
        if (isDateOk)
        {
            SQL += "AND LAST_USED  BETWEEN '" + tboxDateStart.Text + "' AND '" + tboxDateEnd.Text + "'";
        }
        //次數範圍(2者都選了後)
        if (isNumberOk)
        {
            SQL += "AND USED_COUNT BETWEEN " + tboxTagNumStart.Text + " AND " + tboxTagNumEnd.Text + "";
        }
        SQL += " ORDER BY " + this.sortOrder.SelectedValue.ToString() + " " + this.orderBy.SelectedValue.ToString() + " ";
        DataTable TempTable = new DataTable();
        SqlDataAdapter DA = new SqlDataAdapter(SQL, webConfigConnectionString);
        DA.Fill(TempTable);

        if (TempTable.Rows.Count > 0)
        {
            GridViewOfQueryResult.Visible = true;
            GridViewOfQueryResult.DataSource = TempTable;
            GridViewOfQueryResult.DataBind();
        }
        else
        {
            lblNodata.Text = "查無資料";
            GridViewOfQueryResult.Visible = false;
            lblNodata.Visible = true;
        }

    }
}
