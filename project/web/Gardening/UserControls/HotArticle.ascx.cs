using System;
using System.Collections;
using System.Data;
using System.Web.UI;
using System.Data.SqlClient;

public partial class UserControls_HotArticle : UserControl
{
	protected string hotArticle ="";
	private string connString = System.Configuration.ConfigurationManager.ConnectionStrings["ConnString"].ToString();
	private string gardeningConnString = System.Configuration.ConfigurationManager.ConnectionStrings["GardeningconnString"].ToString();
	private SqlConnection myConnection;
	private SqlDataReader myReader;
	protected string articleType="";
	protected string article ="";
	protected int articleNumber = 0;
    protected void Page_Load(object sender, EventArgs e)
    {
		Get_Hot_Article();
    }

    //撈欲顯示的分類id 在每個分類以i寫入id來控制點選顯示的文章
	private void Get_Hot_Article()
	{
		string ss="";
		string st="";
		int count = 0;
		ArrayList number = new ArrayList();
        ArrayList title = new ArrayList(); 
		try
		{
			string sqlString ="";
			myConnection = new SqlConnection(gardeningConnString);
			sqlString = "SELECT typeId , TypeName FROM CATEGORY_TYPE  where Status = 'true'";
			myReader = null;
			myConnection.Open();
			SqlCommand myCommand = new SqlCommand(sqlString, myConnection);
			myReader = myCommand.ExecuteReader();
			
			while(myReader.Read())
			{
				number.Add(myReader["typeId"]);
				title.Add(myReader["TypeName"]);
			}
			myConnection.Close();
			myReader.Close();
		}
		catch(Exception ex)
		{
			Response.Write(ex);
		}
		int i = 0;
		int maxLength=number.ToArray().Length; 
		st= "<table class=\"articleTitle\"  cellpadding=\"0\" cellspacing=\"0\">";
		while(i <= (maxLength))
			{
				if(count == 0)
				{
					st += "<tr><td  style=\"background:url(images/btn_green.png) 0 0 no-repeat;\" onclick=\"Show_Hot_Article("+ i.ToString() +");\" id =\"hot_articleTitle_"+i.ToString()+"\"><a>"+ "全部" + "</a></td></tr>";
					ss += "<table id ='hot_article_" + i.ToString() + "' > <tr><td><ul>" + Hot_Article(0,true)+"</ul></td></tr></table>";
					count += 1;
				}else
				{
					st += "<tr><td onclick=\"Show_Hot_Article("+ i.ToString() +");\" id =\"hot_articleTitle_"+i.ToString()+"\"><a>"+ title[i-1].ToString() + "</a></td></tr>";
					ss += "<table id ='hot_article_" + i.ToString() + "' class=\"hide\" > <tr><td><ul>" + Hot_Article(int.Parse(number[i-1].ToString()),false) +"</ul></td></tr></table>";
				}
			i++;
			}
		if(i != 0)articleNumber=i;
		st +="</table>";
		articleType=st;
		article = ss;
	}
	//撈出知識家發問文章
	private string Hot_Article(int typeId,bool all)
    {
		string html="";
		string sqlString ="";
		try
		{
			myConnection = new SqlConnection(connString);
            sqlString = "SELECT TOP (7) CuDTGeneric.iCUItem, CuDTGeneric.topCat, CuDTGeneric.sTitle, CuDTGeneric.xPostDate ";
			sqlString += "FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem ";
			sqlString += "INNER JOIN GARDENING.dbo.CATEGORY_IN_KNOWLEDGE ON CuDTGeneric.iCUItem = GARDENING.dbo.CATEGORY_IN_KNOWLEDGE.ARTICLEID ";
            if (!all)
            {
                sqlString += "And GARDENING.dbo.CATEGORY_IN_KNOWLEDGE.TYPEID = " + typeId + " ";
            }
			sqlString += "WHERE (iCTUnit = @ictunit) AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) ";
			sqlString += "AND (KnowledgeForum.Status = 'N') ORDER BY CuDTGeneric.xPostDate DESC";
			myReader = null;
			myConnection.Open();
			SqlCommand myCommand = new SqlCommand(sqlString, myConnection);
			myCommand.Parameters.AddWithValue("@ictunit", System.Configuration.ConfigurationSettings.AppSettings["KnowledgeQuestionCtUnitId"].ToString());
			myCommand.Parameters.AddWithValue("@siteid", System.Configuration.ConfigurationSettings.AppSettings["SiteId"].ToString());
			myReader = myCommand.ExecuteReader();
			
			string titleTemp = ""; 
			while(myReader.Read())
			{
				html += "<li><a href=\"/knowledge/knowledge_cp.aspx?ArticleId=" + myReader["iCUItem"] + "&ArticleType=A&CategoryId=" + myReader["topCat"] + "\" target=\"_blank\">";
                titleTemp = Convert.ToString(myReader["sTitle"]);				
				if (titleTemp.Length > 24)
					titleTemp = titleTemp.Substring(0, 24) + "....";										
				html += titleTemp + "</a>";
                html += "<span>" + Convert.ToDateTime(myReader["xPostDate"]).ToShortDateString() + "</span></li>";
			}
			myConnection.Close();
			
			
			myReader.Close();
		}
		catch(Exception ex)
		{
			Response.Write(ex);
		}
		if(html =="") html = "目前此分類沒有文章";
		return html;
    }
}
