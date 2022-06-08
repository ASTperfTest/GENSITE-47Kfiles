using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using GSS.Vitals.COA.Data;

public partial class Story_Detail : System.Web.UI.Page
{
    private string iCTUnit;
    private int ArticleID;
    protected void Page_Load(object sender, EventArgs e)
    {
        iCTUnit = System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnit"].ToString();
        ArticleID = Convert.ToInt32(Request.QueryString["xItem"].ToString());
        string strQueryScript = @"SELECT * FROM CuDTGeneric WHERE iCTUnit = @iCTUnit 
                                    AND fCTUPublic = 'Y' AND iCUItem = @iCUItem
                                  ORDER BY xPostDate DESC";
        try
        {
            using (var reader = SqlHelper.ReturnReader("ConnString", strQueryScript,
                DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnit),
                DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", ArticleID)))
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        labTitle.Text = reader["sTitle"].ToString();
                        labContent.Text = reader["xBody"].ToString();
                        // 加入附件圖片
                        labContent.Text += showPicture(ArticleID.ToString(), "jpg");
                    }
                }
            }

            setRecordURL(ArticleID);
        }
        catch (Exception)
        {
            throw;
        }

    }

    // 上一頁、下一頁的Link
    protected void setRecordURL(int ArticleId)
    {
        string strCreateScript = @"SELECT * INTO #tempCuDtGeneric_Story FROM CuDTGeneric WHERE iCTUnit = @iCTUnit AND fCTUPublic = 'Y' ORDER BY xPostDate DESC ";
        string strBackScript;
        string strNextScript;
        //string strDropScript = " DROP TABLE #tempCuDtGeneric_Story";

        strBackScript = strCreateScript + " SELECT TOP 1 * FROM #tempCuDtGeneric_Story WHERE iCUItem > @iCUItem ORDER BY iCUItem ";
        strNextScript = strCreateScript + " SELECT TOP 1 * FROM #tempCuDtGeneric_Story WHERE iCUItem > @iCUItem ";

        using (var BackRecord = SqlHelper.ReturnReader("ConnString", strBackScript,
            DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnit),
            DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", ArticleId)))
        {
            if (BackRecord.HasRows)
            {
                while (BackRecord.Read())
                {
                    labBack.Text = "<a href=\"/Century/Story_Detail.aspx?xItem=" + BackRecord["iCUItem"].ToString() + "\" class=\"Back\"> 上一篇</a>";
                }
            }
        }
        using (var NextRecord = SqlHelper.ReturnReader("ConnString", strNextScript,
            DbProviderFactories.CreateParameter("ConnString", "@iCTUnit", "@iCTUnit", iCTUnit),
            DbProviderFactories.CreateParameter("ConnString", "@iCUItem", "@iCUItem", ArticleId)))
        {
            if (NextRecord.HasRows)
            {
                while (NextRecord.Read())
                {
                    labNext.Text = "<a href=\"/Century/Story_Detail.aspx?xItem=" + NextRecord["iCUItem"].ToString() + "\" class=\"Next\"> 下一篇</a>";
                }
            }
        }

    }

    /// <summary>
    /// show Attachment(Picture)
    /// </summary>
    /// <param name="iCuItem">CuDTGeneric's PK</param>
    /// <param name="extension">the Filter of extensions</param>
    /// <returns></returns>
    private string showPicture(string iCuItem,string extension)
    {
        string strResult = string.Empty;
        string fiExtension = "." + extension;
        // {0}-title、{1}-Picture URL
        string FileURL = System.Web.Configuration.WebConfigurationManager.AppSettings["WWWURL"].ToString() + "/Public/Attachment/";
        string strTemplate = @"<li><a title={0} href={1}><div class=image><img title={0} alt={0} src={1} />{0}<p/></div></a></li>";
        string strQueryScript = @"SELECT aTitle, NFileName FROM CuDTAttach WHERE xICuItem = @xICuItem AND NFileName LIKE '%' + @extension";

        using (var reader = SqlHelper.ReturnReader("ConnString", strQueryScript,
            DbProviderFactories.CreateParameter("ConnString", "@xICuItem", "@xICuItem", iCuItem),
            DbProviderFactories.CreateParameter("ConnString", "@extension", "@extension", fiExtension)))
        {
            
            if (reader.HasRows)
            {
                // 分段寫只是為了整齊，無其他用意
                strResult = "<DIV class=\"yoxview\">";
                strResult += "<DIV class=\"attachment\">";
                strResult += "<h5>附件</h5>";
                strResult += "<ul>";
                while (reader.Read())
                {
                    strResult += string.Format(strTemplate, reader["aTitle"].ToString(), FileURL + reader["NFileName"].ToString());
                }
                strResult += "</DIV></DIV>";
            }
        }
        
        return strResult;
    }
}