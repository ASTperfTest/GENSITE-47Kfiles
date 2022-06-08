using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Collections.Specialized;

public partial class UserControls_TagCloud : UserControl
{
	protected string tagCloudHtml ="";
	private string gardeningConnString = System.Configuration.ConfigurationManager.ConnectionStrings["GardeningconnString"].ToString();
	private SqlConnection myConnection;
    protected void Page_Load(object sender, EventArgs e)
    {
        GetTagCloud();
    }
    private void GetTagCloud()
	{
        DataTable dtResult = new DataTable();
		try
		{
			string sqlString ="";
			myConnection = new SqlConnection(gardeningConnString);
			myConnection.Open();
            sqlString = "SELECT Top 20 * FROM TAG WHERE USED_COUNT > 10 AND LAST_USED BETWEEN DATEADD(DAY,-7,GETDATE()) AND GETDATE() order by LAST_USED desc";


            SqlDataAdapter SQLDataAdapter = new SqlDataAdapter(sqlString, myConnection);
            SQLDataAdapter.Fill(dtResult);
		}
		catch(Exception ex)
		{
			Response.Write(ex);
		}finally
		{
			if (myConnection.State == ConnectionState.Open)
            {
				myConnection.Close();
            }
		}
        tagCloudHtml = LoadData(dtResult);
	}

    public string LoadData(DataTable table)
    {
        int max = 0;
        string taghtml = "";
        foreach (DataRow row in table.Rows)
        {
            if (int.Parse(row["USED_COUNT"].ToString()) > max) max = int.Parse(row["USED_COUNT"].ToString());
        }
        foreach (DataRow row in table.Rows)
        {
            int tagcss = 0;
            tagcss = GetTagLevel(row["DISPLAY_NAME"].ToString(), int.Parse(row["USED_COUNT"].ToString()), max,table.Rows.Count);
            taghtml += "<a class=\"TagLv_" + tagcss.ToString() + "\" herf=\"\"  onclick=\"" + "javascript:document.SearchForm.Keyword.value ='" + row["DISPLAY_NAME"].ToString() + "'; checkSearchForm(0)" + "\" onmouseover=\"this.style.color='#0083E5';\" onmouseout=\"this.style.color='#0B5891';\">" + row["DISPLAY_NAME"].ToString() + "</a>";
        }
        return taghtml;
    }


    private int GetTagLevel(string TagName, int count, int max,int tagCount)
    {
        int TopLevel = 7;

        int range = 0;
        if (tagCount <= TopLevel)
        {
            range = TopLevel;
        }
        else
        {
            range = max / TopLevel;
        }

        int result = TopLevel;

        while (result > 1)
        {
            if (count > range * result)
            {
                break;
            }
            result--;
        }
        return result;
    }
}
