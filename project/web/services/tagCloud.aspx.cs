using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using System.Collections.Specialized;

public partial class services_tagCloud : System.Web.UI.Page
{
    private string ConnString = System.Configuration.ConfigurationManager.ConnectionStrings["GSSConnString"].ToString();
    protected void Page_Load(object sender, EventArgs e)
    {
        string cmd = WebUtility.GetStringParameter("cmd", string.Empty).ToLower();
        if (cmd == string.Empty)
        {
            return;
        }
        else
        {
            DispatchMethod(cmd);
        }
    }

    private void DispatchMethod(string MethodName)
    {
        switch (MethodName)
        {
            case "gettagcloud":
                GetTagCloud();
                break;
            case "addtag":
                AddTag();
                break;
            default:
                break;
        }
    }

    private void GetTagCloud()
    {
        using (SqlConnection conn = new SqlConnection(ConnString))
        {
            DataTable dtResult = new DataTable();
            try
            {
                conn.Open();
                string SQLstringt = "SELECT Top 20 * FROM TAG WHERE USED_COUNT > 10 AND LAST_USED BETWEEN DATEADD(DAY,-7,GETDATE()) AND GETDATE() order by LAST_USED desc";

                SqlDataAdapter SQLDataAdapter = new SqlDataAdapter(SQLstringt, conn);
                SQLDataAdapter.Fill(dtResult);
                if (dtResult.Rows.Count >= 1)
                {
                    WebUtility.WriteAjaxResult(true, LoadData(dtResult));
                }
                else
                {
                    WebUtility.WriteAjaxResult(true, "");
                }
            }
            catch (Exception ex)
            {
                //目前都會有connection pool滿了的問題，exception先不show出了...
                //Response.Write(ex);
            }finally
		{
			if (conn.State == ConnectionState.Open)
            {
				conn.Close();
            }
		}
        }
    }

    private void AddTag()
    {
	    NameValueCollection RequestUrl = HttpUtility.ParseQueryString(Request["tag"], Encoding.GetEncoding("utf-8"));
        string tagname = HttpUtility.HtmlDecode(WebUtility.GetStringParameter("tag", string.Empty));
        if (RequestUrl["Keyword"] != "")
        {
            tagname = RequestUrl["Keyword"];
        }
       string[] tagNames = null;
	   char[] delimiterChars = { ' ', ',', '.', ':', '\t',';','　','、' };
       tagNames = tagname.Split(delimiterChars);
       int i = 0;
       while (i < tagNames.Length)
       {
	       if (TagExist(tagNames[i].ToString()))
            {
                UpdateTag(tagNames[i].ToString());
            }
            else
            {
                InsertTag(tagNames[i].ToString());  
            }
           i++;
       }
       WebUtility.WriteAjaxResult(true, "", null);
    }

    public string LoadData(DataTable table)
    {
        int max = 0;
        string taghtml = "";
        foreach (DataRow row in table.Rows)
        {
            if(int.Parse(row["USED_COUNT"].ToString()) > max) max =int.Parse(row["USED_COUNT"].ToString());
        }
        foreach (DataRow row in table.Rows)
        {
            int tagcss = 0;
			tagcss = GetTagLevel(row["DISPLAY_NAME"].ToString(), int.Parse(row["USED_COUNT"].ToString()), max,table.Rows.Count);
            taghtml += "<a class=\"TagLv_" + tagcss.ToString() + "\" herf=\"\"  onclick=\"" +"javascript:document.SearchForm.Keyword.value ='" +row["DISPLAY_NAME"].ToString()+"'; checkSearchForm(0)"+"\" onmouseover=\"this.style.color='#0083E5';\" onmouseout=\"this.style.color='#0B5891';\">" + row["DISPLAY_NAME"].ToString() + "</a>";
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
	
	private bool TagExist(string name)
    {
        SqlConnection conn = new SqlConnection(ConnString);
        SqlDataReader reader = null;
        DataTable dtResult = new DataTable();
        try
        {
            SqlCommand cmd = new SqlCommand("SELECT * FROM TAG Where DISPLAY_NAME = @DISPLAY_NAME",conn);

            cmd.CommandType = CommandType.Text;
            SqlParameter param = new SqlParameter();
            param.ParameterName = "@DISPLAY_NAME";
            param.Value = name.Trim();
            cmd.Parameters.Add(param);
            conn.Open();
            reader = cmd.ExecuteReader();
            
            dtResult.Load(reader);
            if (dtResult.Rows.Count >= 1) return true;
        }
        catch (Exception ex)
        {
            Response.Write(ex);
        }finally
		{
			if (conn.State == ConnectionState.Open)
            {
				conn.Close();
            }
		}
        return false;
    }

    private void UpdateTag(string name)
    {
         SqlConnection conn = new SqlConnection(ConnString);
        SqlDataReader reader = null;
        DataTable dtResult = new DataTable();
        try
        {
            SqlCommand cmd = new SqlCommand("update TAG SET USED_COUNT =  USED_COUNT + 1,LAST_USED = @LAST_USED  WHERE DISPLAY_NAME = @DISPLAY_NAME", conn);
            cmd.CommandType = CommandType.Text;
            SqlParameter param = new SqlParameter();
            param.ParameterName = "@DISPLAY_NAME";
            param.Value = name;
            cmd.Parameters.Add(param);
            SqlParameter param1 = new SqlParameter();
            param1.ParameterName = "@LAST_USED";
            param1.Value = DateTime.Now;
            cmd.Parameters.Add(param1);
            conn.Open();
            reader = cmd.ExecuteReader();
            
            dtResult.Load(reader);
        }
        catch (Exception ex)
        {
            Response.Write(ex);
        }finally
		{
			if (conn.State == ConnectionState.Open)
            {
				conn.Close();
            }
		}
    }

    private void InsertTag(string name)
    {
        SqlConnection conn = new SqlConnection(ConnString);
        SqlDataReader reader = null;
        DataTable dtResult = new DataTable();
        try
        {
            SqlCommand cmd = new SqlCommand("INSERT INTO TAG (DISPLAY_NAME, USED_COUNT, LAST_USED)VALUES (@DISPLAY_NAME, 1, @LAST_USED)", conn);
			cmd.CommandType = CommandType.Text;
            SqlParameter param = new SqlParameter();
            param.ParameterName = "@DISPLAY_NAME";
			if (name.Length >= 225)
			{
                name = name.Substring(0,224);
			}
            param.Value = name;
            cmd.Parameters.Add(param);
            SqlParameter param1 = new SqlParameter();
            param1.ParameterName = "@LAST_USED";
            param1.Value = DateTime.Now;
            cmd.Parameters.Add(param1);
            conn.Open();
            reader = cmd.ExecuteReader();

            dtResult.Load(reader);
        }
        catch (Exception ex)
        {
			//derek 2009/12/12
			//遇到tag 重複情況,吃掉
            //Response.Write(ex);
        }finally
		{
			if (conn.State == ConnectionState.Open)
            {
				conn.Close();
            }
		}
    }
}
