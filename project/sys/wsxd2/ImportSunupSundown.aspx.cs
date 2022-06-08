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

using System.IO;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Transactions;
using System.Globalization;


/// <summary>
/// 離線編輯Excel並提供線上上傳
/// </summary>
public partial class OfflineExcel : System.Web.UI.Page
{
    string gUploadFileName ;

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    /// <summary>
    /// 下載Excel範本
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
    protected void Download_Click(object sender, EventArgs e)
    {
        Page.Response.ContentType = "application/vnd.ms-excel";//內容類型
        Page.Response.AddHeader("content-disposition", "attachment; filename=" + "TemplateSunData.xls");
        FileStream tDownFile = new FileStream(System.Web.HttpContext.Current.Server.MapPath("./OfflineDown/TemplateSunData.xls"), FileMode.Open);
        long tFileSize;
        tFileSize = tDownFile.Length;
        byte[] tContent = new byte[(int)tFileSize];
        tDownFile.Read(tContent, 0, (int)tDownFile.Length);
        tDownFile.Close();
        Page.Response.BinaryWrite(tContent);
    } 

    /// <summary>
    /// 將Excel上傳到server端
    /// </summary>
    private void saveUploadFile()
    {
        HttpPostedFile tUploadFile = FileUploadExcel.PostedFile;
        int tFileLength = tUploadFile.ContentLength;
        byte[] tFileByte = new byte[tFileLength];
        tUploadFile.InputStream.Read(tFileByte, 0, tFileLength);

        FileStream tNewfile = new FileStream(System.Web.HttpContext.Current.Server.MapPath("./OfflineUp/") + DateTime.Now.ToString("yyyyMMddhhmm") + "_upload.xls", FileMode.Create);
        tNewfile.Write(tFileByte, 0, tFileByte.Length);
        tNewfile.Close();

        //這是一個全域變數，記錄excel的上傳路徑  
        gUploadFileName = tNewfile.Name;
    }

    /// <summary>
    /// 讀取Excel內容，並儲存至 DataTable
    /// </summary>
    /// <returns>含有Excel內容的DataTable</returns>
    private DataTable readDataFromXls()
    {
       
        OleDbConnection tConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + gUploadFileName + ";Extended Properties=Excel 8.0;");
       
        try
        {
            #region 日出日落

            tConn.Open();   //開啟excel路徑  
            DataSet tDs = new DataSet();
            OleDbDataAdapter tDa = new OleDbDataAdapter("Select * From [template$] ", tConn);//...From [sheet name+$]
            tDa.Fill(tDs, "SunupSundown"); //將excel內的sheet讀入dataset中  ；預設第一列為 table的 column name
            DataTable tDt = tDs.Tables["SunupSundown"]; //自訂 table name，並且一個DataSet可存放多張table
            return tDt;

            #endregion


        }
        catch(Exception ex)
        {
            if (ex.Message.Contains("$' 不是有效的名稱"))
            {
                string tErroralert1 = "<script>alert('找不到Excel中符合的Sheet。')</script>";
                ClientScript.RegisterStartupScript(typeof(OfflineExcel), "formaterror", tErroralert1);
                return null;
            }
            System.IO.FileInfo tDeletfile = new FileInfo(gUploadFileName);
            tConn.Close();//by ivy 需要先關掉對於Excel的開啟，才能執行下行的刪除。
            tDeletfile.Delete();
            string tErroralert = "<script>alert('上傳Excel格式錯誤')</script>";
            ClientScript.RegisterStartupScript(typeof(OfflineExcel), "formaterror", tErroralert);
            return null;
        }
        finally
        {
            System.IO.FileInfo tDeletfile = new FileInfo(gUploadFileName);
            tConn.Close();
            tDeletfile.Delete();
        }
    }

    /// <summary>
    /// 新增日出日落 Excel資料於資料庫中
    /// </summary>
    /// <param name="sunUpDownData">日出日落 Excel</param>
    private void insertSunUpAndDownDataToDB(DataTable sunUpDownData)
    {
        using (TransactionScope scope = new TransactionScope())
        {
            using (SqlConnection tSqlConn = new SqlConnection())
            {
                string webConfigConnectionString = 
                    System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

                tSqlConn.ConnectionString = webConfigConnectionString;

                try
                {
                    tSqlConn.Open();

                    if (sunUpDownData.Rows.Count > 0)
                    {
                        #region Delete

                        //讀取年份及區域。資料位於第一列
                        string year = sunUpDownData.Rows[0][0].ToString();
                        string region = sunUpDownData.Rows[0][1].ToString();

                        //判斷該年份資料是否已存在。存在就將之刪除。
                        if (querySunDataCountsByYear(tSqlConn, year, region) > 0)
                        {
														//刪除附表
														string deleteSubCommand =
                                "delete from CuDTx7 where giCuItem in (select iCUItem from CuDTGeneric where iBaseDSD='7' and iCTUnit='303' and xPostDate like '%' + @year + '%' and (xKeyword is null or xKeyword = @xKeyword ))";
                            SqlCommand delSubCmd = new SqlCommand(deleteSubCommand, tSqlConn);
                            delSubCmd.Parameters.AddWithValue("year", year);
                            delSubCmd.Parameters.AddWithValue("xKeyword", region);
                            delSubCmd.ExecuteNonQuery();
							
														//刪除主表
														string deleteMainCommand =
                                "delete CuDTGeneric where iBaseDSD='7' and iCTUnit='303' and xPostDate like '%' + @year + '%' and (xKeyword is null or xKeyword = @xKeyword )";
                            SqlCommand delMainCmd = new SqlCommand(deleteMainCommand, tSqlConn);
                            delMainCmd.Parameters.AddWithValue("year", year);
                            delMainCmd.Parameters.AddWithValue("xKeyword", region);
                            delMainCmd.ExecuteNonQuery();
                        }
                        
                        #endregion

                        //新增日出日落資料，從第二列開始

                        int firstMonth = 1;
                        int lastMonth = 12;
                        
                        string insertCommand = @"INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]
                                                               ([iBaseDSD]
                                                               ,[iCTUnit]          
                                                               ,[sTitle]
                                                               ,[iEditor]
                                                               ,[iDept]
                                                               ,[siteId]
                                                               ,[fCTUPublic]
                                                               ,[xKeyword]          
                                                               ,[xPostDate]        
                                                               ,[xBody])
                                                         VALUES
                                                               (7
                                                               ,303          
                                                               ,@sTitle
                                                               ,'gss'
                                                               ,'0'
                                                               ,'1'
                                                               ,'Y'
                                                               ,@region
                                                               ,@xPostDate        
                                                               ,@xBody)";

                        SqlCommand tSqlCmd = new SqlCommand(insertCommand, tSqlConn);
                        
                        string stringPostDate = string.Empty;
                        DateTime postDate;
                       
                        for (int nowMonth = firstMonth; nowMonth <= lastMonth; nowMonth++)
                        {
                            int indexSunUpOfMonth = nowMonth + nowMonth + 1;
                            int indexSunDownOfMonth = indexSunUpOfMonth + 1;

                            for (int i = 1; i < sunUpDownData.Rows.Count; i++)//sunUpDownData.Rows.Count
                            {
                                if (sunUpDownData.Rows[i][indexSunUpOfMonth].ToString() != "")
                                {
                                	  stringPostDate = year + "/" + nowMonth + "/" + sunUpDownData.Rows[i][2].ToString();

                                    if (!DateTime.TryParse(stringPostDate, out postDate))
                                    {
                                        //ClientScript.RegisterStartupScript(Page.GetType(), "dateValidate", "alert('" + stringPostDate + "非正確日期！');", true);
                                        throw new Exception(stringPostDate + "非正確日期！");
                                    }
                                    else
                                    {
		                                    string sunUp = ((DateTime)sunUpDownData.Rows[i][indexSunUpOfMonth]).ToString("HH:mm");
		                                    string sunDown = ((DateTime)sunUpDownData.Rows[i][indexSunDownOfMonth]).ToString("HH:mm");
		                                    string sunUpAndDown = "日出:" + sunUp + " 日落:" + sunDown;
		
		                                    #region 取得農曆年
		                                    string lunisolarDate = string.Empty;
		
		                                    TaiwanLunisolarCalendar cal = new TaiwanLunisolarCalendar();
		                                    
		                                    String CelestialStem = "0甲乙丙丁戊己庚辛壬癸";
		                                    String TerrestrialBranch = "0子丑寅卯辰巳午未申酉戌亥";
		
		                                    int lun60Year = cal.GetSexagenaryYear(postDate);
		                                    int CelestialStemYear = cal.GetCelestialStem(lun60Year);
		                                    int TerrestrialBranchYear = cal.GetTerrestrialBranch(lun60Year);
		
		                                    int lunMonth = cal.GetMonth(postDate);
		                                    int leapMonth = cal.GetLeapMonth(cal.GetYear(postDate));
		
		                                    if (leapMonth > 0 && lunMonth >= leapMonth)
		                                    {
		                                        lunMonth -= 1;
		                                    }
		                                    int lunDay = cal.GetDayOfMonth(postDate);
		
		                                    lunisolarDate = (String.Format("農曆{0}年{1}月{2}日",
		                                       CelestialStem[CelestialStemYear].ToString() + TerrestrialBranch[TerrestrialBranchYear].ToString(), lunMonth, lunDay));
		
		                                    #endregion
		
		                                    #region Insert
		
		                                    tSqlCmd.Parameters.AddWithValue("sTitle", lunisolarDate);
		                                    tSqlCmd.Parameters.AddWithValue("region", region);
		                                    tSqlCmd.Parameters.AddWithValue("xPostDate", postDate);
		                                    tSqlCmd.Parameters.AddWithValue("xBody", (region + " " +sunUpAndDown));
		
		                                    tSqlCmd.ExecuteNonQuery();
		                                    tSqlCmd.Parameters.Clear();
		                                    stringPostDate = string.Empty;
		                                    #endregion
                                  	}
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        string tErroralert = "<script>alert('匯入資料成功!!')</script>";
						insertCuDTx7(tSqlConn,year,region);//新增CuDTx7資料表資料
                        ClientScript.RegisterStartupScript(typeof(OfflineExcel), "formaterror", tErroralert);

                        scope.Complete();

                    }//end if
                }
                catch (Exception ex)
                {
                    string tErroralert = "<script>alert('匯入資料失敗!!"+ex.Message.ToString()+"' )</script>";
                    ClientScript.RegisterStartupScript(typeof(OfflineExcel), "formaterror", tErroralert);
                }
                
            }
        }
    }

    /// <summary>
    /// 查詢該年度下的日出日落資料筆數
    /// </summary>
    /// <param name="year">該年度</param>
    /// <param name="region">該區域</param>
    /// <returns>查詢資料筆數</returns>
    private int querySunDataCountsByYear(SqlConnection tSqlConn, string year, string region)
    {
        if (tSqlConn.State == ConnectionState.Closed)
        {
            tSqlConn.Open();
        }

        string selectCommand = "select count(1) from CuDTGeneric where iBaseDSD='7' and iCTUnit='303' and xPostDate like '%' + @year + '%'  and (xKeyword is null or xKeyword = @xKeyword )";

        SqlCommand tCmd = new SqlCommand(selectCommand, tSqlConn);
        tCmd.Parameters.AddWithValue("year", year);
        tCmd.Parameters.AddWithValue("xKeyword", region);

        int sunDataCounts = (int)tCmd.ExecuteScalar();

        return sunDataCounts;
    }
	
	private void insertCuDTx7(SqlConnection tSqlConn,string year, string region)
	{
		string insCommand = @"INSERT INTO CuDTx7
							SELECT iCUItem,null FROM CuDtGeneric
							WHERE ictunit=303 
							AND xKeyword=@xKeyword
							AND xPostDate like '%' + @year + '%'";
							
		SqlCommand tCmd = new SqlCommand(insCommand, tSqlConn);
        tCmd.Parameters.AddWithValue("xKeyword", region);
		tCmd.Parameters.AddWithValue("year", year);
		
		tCmd.ExecuteNonQuery();
	}

    /// <summary>
    /// 匯入Excel資料到資料庫中
    /// </summary>
    protected void ButtonUpload_Click(object sender, EventArgs e)
    {
        if (FileUploadExcel.HasFile)
        {
            saveUploadFile();
            DataTable dt = readDataFromXls();
            insertSunUpAndDownDataToDB(dt);
        }
        else
        {
            string tErroralert = "<script>alert('請選擇上傳檔案。')</script>";
            ClientScript.RegisterStartupScript(typeof(OfflineExcel), "formaterror", tErroralert);
            //ScriptManager.RegisterStartupScript(this.Page, typeof(OfflineExcel), "formaterror", tErroralert, false);//true會出現「語法錯誤」
        }
        
    } 
}
