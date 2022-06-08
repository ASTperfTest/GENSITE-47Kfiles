using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.IO;
using System.Xml;
using System.Data.SqlClient;
using System.Collections;
using System.Data;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;

public partial class SearchResult : System.Web.UI.Page
{
    private IGardeningService gardeningService;
    protected String searchResultStr = "";
    protected String searchTopicStr = "";
    protected String Keyword = "";
    private String myXdUrl = System.Web.Configuration.WebConfigurationManager.AppSettings["myXDURL"] + "wsxd2/xdSPHorticulture.asp";
    protected string url = System.Web.Configuration.WebConfigurationManager.AppSettings["myURL"];
    protected String knowledgeUrl ="";
    private string gardeningConnString = System.Configuration.ConfigurationManager.ConnectionStrings["GardeningconnString"].ToString();
    protected String param = "";
    protected void Page_Load(object sender, EventArgs e)
    {
		
        Keyword = HttpUtility.HtmlEncode(WebUtility.GetStringParameter("Keyword", string.Empty));
        if (Keyword == "")
        {
            Response.Redirect("/Gardening/index.aspx");
        }

        if(!IsPostBack)
        {
            searchTopicStr = GetTopicSearch();
            searchResultStr = GetSearch();
            AddCloud();
        }
    }

    //從asp頁面撈回知識家xml
    private String GetSearch()
    {
        param = "?xdURL=Search/HSearchResultList.asp&mp=1&Keyword=" + HttpUtility.UrlEncode(Keyword) + "&FromKnowledgeHome=1&memID=&gstyle=";
	    String ss="<table width=\"100%\">";
        String st = "";
		
    	try
        {
			XmlNodeList nodeList;
			XmlDocument xmlDoc = new XmlDocument();
            XmlReader reader = XmlReader.Create(myXdUrl + param);
	    	xmlDoc.Load(reader);
	    	XmlNode root = xmlDoc.DocumentElement;
	    	nodeList = root.SelectNodes("//pHTML/doclist");
	    	foreach ( XmlNode doclist in nodeList)
            {
                st += WriteKnowledgeHtml(doclist);
            }
        }
    	catch(Exception ex)
        {
	    	//Response.Write(ex.ToString());
            Response.Redirect("/Gardening/index.aspx");
        }

        if (st == "")
        {
            st = "<tr><td>無相關資料</td></tr>";
        }
        ss += st;
        ss += "<tr><td colSpan=\"2\" align=\"right\"><div class=\"more\" align=\"right\"><a  align=\"right\" onclick=\"moreKnowledge( )\">more...</a></div></td></tr></table>";
		
        return ss;
    }
    //寫出知識家查詢結果
    private String WriteKnowledgeHtml(XmlNode doclist)
    {
        knowledgeUrl = url + "/knowledge/knowledge_cp.aspx";
	    String stringTemp="";
	    for(int i = 0; i < doclist.ChildNodes.Count;i++)
        {
	    	if (i == 0 )
            {
	    		stringTemp += "<tr><td>";
		    	stringTemp += "<a href=\"" + knowledgeUrl + "?ArticleId=" + doclist.ChildNodes[i].InnerText + "&ArticleType=A&CategoryId=E&kpi=0\" target=\"_blank\">";
		    }
            else if (i == 1)
            {
		    	stringTemp += doclist.ChildNodes[i].InnerText;
		    	stringTemp += "</a>";
            }
		    else if( i == 2)
            {
                stringTemp += "<span>" + doclist.ChildNodes[i].InnerText + "</span></td>";
            }
	    }
	    stringTemp += "</tr>";
        return stringTemp;
    }

    //搜尋日誌與作品本頁只帶出10筆
    private string GetTopicSearch()
    {
        string ss = "<table width=\"100%\">";
        gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        IList topicList = gardeningService.GetTopicsSearchResult(Keyword, 0, 10);
        if (topicList.Count > 0)
        {
            foreach (Topic topic in topicList)
            {
                ss += "<tr><td>";
                ss += "<a href=\"" + url + "/Gardening/entrylist.aspx?topicid=" + topic.TopicId + "\" >";
                ss += topic.Title + "</a>";
                ss += "<span>" + Convert.ToDateTime(topic.ModifyDateTime).ToShortDateString() + "</span>";
                ss += "</td></tr>";
            }
        }
        else
        {
            ss += "<tr><td>作品沒有相關內容</td></tr>";
        }
        ss += "<tr><td colSpan=\"2\"><div class=\"more\" align=\"right\"><a href=\"/Gardening/topiclist.aspx?Keyword=" + Keyword + "\">more...</a></div></td></tr></table>";
        return ss;
    }

    //將搜尋的keyword紀錄標籤雲次數
    private void AddCloud()
    {
        string tagname = Keyword;
        string[] tagNames = null;
        char[] delimiterChars = { ' ', ',', '.', ':', '\t',';','　','、' };
        tagNames = tagname.Split(delimiterChars);
        int i = 0;
        while (i < tagNames.Length)
        {
            if (TagExist(tagNames[i].ToString()))
            {
                UpdateTag(tagNames[i].ToString(),"update");
            }
            else
            {
                UpdateTag(tagNames[i].ToString(),"insert");
            }
            i++;
        }
    }

    //判斷是否已經有此tag
    private bool TagExist(string name)
    {
        SqlConnection cnn = new SqlConnection(gardeningConnString);
        SqlDataReader reader = null;
        DataTable dtResult = new DataTable();
        try
        {
            SqlCommand cmd = new SqlCommand("SELECT * FROM TAG Where DISPLAY_NAME = @DISPLAY_NAME", cnn);
            cmd.CommandType = CommandType.Text;
            SqlParameter param = new SqlParameter();
            param.ParameterName = "@DISPLAY_NAME";
            param.Value = name.Trim();
            cmd.Parameters.Add(param);
            cnn.Open();
            reader = cmd.ExecuteReader();

            dtResult.Load(reader);
            if (dtResult.Rows.Count >= 1) return true;
        }
        catch (Exception ex)
        {
            Response.Write(ex);
        }finally
		{
			if (cnn.State == ConnectionState.Open)
            {
				cnn.Close();
            }
		}
        return false;
    }

    //更新標籤雲count
    private void UpdateTag(string name,string action)
    {
        SqlConnection cnn = new SqlConnection(gardeningConnString);
        SqlDataReader reader = null;
        DataTable dtResult = new DataTable();
        string sqlstring = "";
        if (action == "update")
        {
            sqlstring = "update TAG SET USED_COUNT =  USED_COUNT + 1,LAST_USED = @LAST_USED  WHERE DISPLAY_NAME = @DISPLAY_NAME";
        }
        else if (action == "insert")
        {
            sqlstring = "INSERT INTO TAG (DISPLAY_NAME, USED_COUNT, LAST_USED)VALUES (@DISPLAY_NAME, 1, @LAST_USED)";
        }

        try
        {

            SqlCommand cmd = new SqlCommand(sqlstring, cnn);
            cmd.CommandType = CommandType.Text;
            SqlParameter param = new SqlParameter();
            param.ParameterName = "@DISPLAY_NAME";
            if (name.Length >= 225)
            {
                name = name.Substring(0, 224);
            }
            param.Value = name;
            cmd.Parameters.Add(param);
            SqlParameter param1 = new SqlParameter();
            param1.ParameterName = "@LAST_USED";
            param1.Value = DateTime.Now;
            cmd.Parameters.Add(param1);
            cnn.Open();
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
			if (cnn.State == ConnectionState.Open)
            {
				cnn.Close();
            }
		}
    }
}
