using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

namespace 監看系統
{
    public partial class _Default : System.Web.UI.Page
    {
        DataTable g_table = new DataTable();
        DataView dataView = null;
        protected void Page_Load(object sender, EventArgs e)
        {
            
            if (!Page.IsPostBack )
            {
              
            }

           
            BindGrid();
            
         //   g_table.Rows.Add(g_row);
           


        }
        private string strConnString()
        {
            return System.Configuration.ConfigurationSettings.AppSettings["connectionstring"];

        }

        public void ExportExcel()
        {


            using (SqlConnection connection = new SqlConnection(strConnString()))
            {
                connection.Open();
                
                //SqlConnection connection = this.openConn();
                if (connection == null) return;
                string SQL = "SELECT [UID], [REALNAME], [NICKNAME], [TEL], [EMAIL], [ADDRESS], [TOTALSTAR], [TOTALSCORE], [PLANT_A_SCORE], [PLANT_A_STAR], [PLANT_A_PSEUDO_STAR], [PLANT_A_CURRENT_SCORE], [PLANT_B_SCORE], [PLANT_B_STAR], [PLANT_B_PSEUDO_STAR], [PLANT_D_SCORE], [PLANT_C_CURRENT_SCORE], [PLANT_D_STAR], [PLANT_D_PSEUDO_STAR], [PLANT_D_CURRENT_SCORE], [PLANT_B_CURRENT_SCORE], [PLANT_C_SCORE], [PLANT_C_STAR], [PLANT_C_PSEUDO_STAR] FROM [LBG_SCORE_SUMMARY] ";
                if ( filterExpression() != "" ) SQL = SQL + " where " + filterExpression();
                SQL = SQL + " ORDER BY [TOTALSTAR] DESC, [TOTALSCORE] DESC";

                this.EnableViewState = false;
                string fileName =  "RANKING_REPORT-"+DateTime.Now.ToString ("yyyyMMdd_hhmmss") +".xls" ; // DateTime.Now.Year.ToString ()+"-"+DateTime.Now.Month.ToString ()+"-"+DateTime.Now.Day.ToString ()+"-"+DateTime.Now.Hour.ToString()+"-"+DateTime.Now.Minute.ToString ()+"-"+DateTime.Now.Second.ToString () + ".xls"
                string headerStr =String.Format ( "attachment;filename={0}", fileName );
                
                HttpResponse resp;
                resp = Page.Response;
                Response.Clear();
                Response.AddHeader("content-disposition", headerStr);
    
                Response.Buffer = true;
                Response.Charset = "BIG5";

                Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5");
            
                //Response.ContentType = "application/ms-excel";//设置输出文件类型为excel文件。 
                Response.ContentType = "application/vnd.ms-excel";

                string colHeaders= "", ls_item="";
                int i=0;
                //取得数据表各列标题，各标题之间以\t分割，最后一个列标题后加回车符
                colHeaders = "[UID]\t [REALNAME]\t [NICKNAME]\t [TEL]\t [EMAIL]\t [ADDRESS]\t [TOTALSTAR]\t [TOTALSCORE]\t [PLANT_A_SCORE]\t [PLANT_A_STAR]\t [PLANT_A_PSEUDO_STAR]\t [PLANT_A_CURRENT_SCORE]\t [PLANT_B_SCORE]\t [PLANT_B_STAR]\t [PLANT_B_PSEUDO_STAR]\t [PLANT_D_SCORE]\t [PLANT_C_CURRENT_SCORE]\t [PLANT_D_STAR]\t [PLANT_D_PSEUDO_STAR]\t [PLANT_D_CURRENT_SCORE]\t [PLANT_B_CURRENT_SCORE]\t [PLANT_C_SCORE]\t [PLANT_C_STAR]\t [PLANT_C_PSEUDO_STAR] \n";
                //向HTTP输出流中写入取得的数据信息
                Response.Write(colHeaders); 


                SqlCommand cmd = new SqlCommand();
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlDataReader result = cmd.ExecuteReader();
                if (!result.Read()) return ;
                try
                {
                    int count = 0;
                    do
                    {
                        //在当前行中，逐列获得数据，数据之间以\t分割，结束时加回车符\n
                        for(i=0;i <= result.FieldCount  -1 ;i++)
                            ls_item +=result[i].ToString() + "\t";     
                        ls_item += "\n";
                        //当前行数据写入HTTP输出流，并且置空ls_item以便下行数据    
                        Response.Write(ls_item);
                        ls_item="";
                    }
                    while (result.Read());
                }
                catch
                {
                    return ;
                }
                finally
                {
                    cmd.Cancel();
                    result.Close();
                }

                Response.End();
                this.EnableViewState = true;
            }











































//            SqlDataSource.FilterExpression = filterExpression();
//            SqlDataSource.DataBind();
            
//            DataView dv = (DataView)(SqlDataSource.Select(DataSourceSelectArguments.Empty));

//            DataSet newDS = new DataSet();

//            DataTable g_table = dv.Table.Clone();

//            foreach (DataRowView drv in dv)

//                g_table.ImportRow(drv.Row);

////            newDS.Tables.Add(dt);

//            this.EnableViewState = false;
//            string fileName =  "RANKING_REPORT-"+DateTime.Now.ToString ("yyyyMMdd_hhmmss") +".xls" ; // DateTime.Now.Year.ToString ()+"-"+DateTime.Now.Month.ToString ()+"-"+DateTime.Now.Day.ToString ()+"-"+DateTime.Now.Hour.ToString()+"-"+DateTime.Now.Minute.ToString ()+"-"+DateTime.Now.Second.ToString () + ".xls"
//            string headerStr =String.Format ( "attachment;filename={0}", fileName );
//            //string headerStr ="attachment;filename=FileName.xls";
//            HttpResponse resp;
//            resp = Page.Response;
//            Response.Clear();
//            Response.AddHeader("content-disposition", headerStr);

//            Response.Buffer = true;
//            Response.Charset = "BIG5";

//            Response.ContentEncoding = System.Text.Encoding.GetEncoding("BIG5");
            
//            //Response.ContentType = "application/ms-excel";//设置输出文件类型为excel文件。 
//            Response.ContentType = "application/vnd.ms-excel";

//            string colHeaders= "", ls_item="";
//            int i=0;
            
//            //取得数据表各列标题，各标题之间以\t分割，最后一个列标题后加回车符
//            for(i=0;i<= g_table.Columns.Count-1 ; i++)
//            {
//                colHeaders+=g_table.Columns[i].Caption.ToString()+"\t";
//            }
//            colHeaders +="\n";   
//            //向HTTP输出流中写入取得的数据信息
//            Response.Write(colHeaders); 
//            //逐行处理数据  
//            foreach (DataRow row in dv.Table.Rows)
//            {
//                //在当前行中，逐列获得数据，数据之间以\t分割，结束时加回车符\n
//                for(i=0;i <= row.ItemArray.Length -1 ;i++)
//                    ls_item +=row[i].ToString() + "\t";     
//                ls_item += "\n";
//                //当前行数据写入HTTP输出流，并且置空ls_item以便下行数据    
//                Response.Write(ls_item);
//                ls_item="";
//            }
//            Response.End();
//            this.EnableViewState = true;
        
        }

        public string filterExpression()
        {

            string caluse1 = "";
            string caluse2 = "";
            string caluse3 = "";
            string caluse3_1 = "";
            string caluse3_2 = "";
            string filterExp = "";
            if (this.txtName.Text != "") caluse1 = "REALNAME like '%" + this.txtName.Text + "%'";
            if (this.txtNickName.Text != "") caluse2 = "NICKNAME like '%" + this.txtNickName.Text + "%'";
            int ScoreH = 0;
            int ScoreL = 0;
            try
            {
                ScoreH = System.Convert.ToInt32(this.txtScoreHighBound.Text);
                if (this.txtScoreHighBound.Text != "") caluse3_1 = "TOTALSCORE < " + this.txtScoreHighBound.Text;
            }
            catch
            {

                caluse3_1 = "";
            }

            try
            {
                ScoreL = System.Convert.ToInt32(this.txtScoreLowBound.Text);
                if (this.txtScoreLowBound.Text != "") caluse3_2 = "TOTALSCORE >= " + txtScoreLowBound.Text;
            }
            catch
            {
                caluse3_2 = "";
            }

            if ((caluse3_1 != "") && (caluse3_2 == "")) caluse3 = caluse3_1;
            if ((caluse3_1 == "") && (caluse3_2 != "")) caluse3 = caluse3_2;

            //若小 > 大 == > 字串對調
            if ((caluse3_1 != "") && (caluse3_2 != ""))
            {
                if (ScoreL > ScoreH)
                {
                    caluse3_1 = "TOTALSCORE < " + txtScoreLowBound.Text;
                    caluse3_2 = "TOTALSCORE >= " + this.txtScoreHighBound.Text;
                }
            }

            if ((caluse3_1 != "") && (caluse3_2 != "")) caluse3 = "(" + caluse3_1 + " AND " + caluse3_2 + ")";

            //組合Filter String
            if (caluse1 != "") filterExp = caluse1;
            if ((caluse2 != "") && (filterExp != "")) filterExp = filterExp + " and " + caluse2;
            if ((caluse2 != "") && (filterExp == "")) filterExp = caluse2;
            if ((caluse3 != "") && (filterExp != "")) filterExp = filterExp + " and " + caluse3;
            if ((caluse3 != "") && (filterExp == "")) filterExp = caluse3;

            return filterExp;

        }

        public void BindGrid()
        {


            Label5.Text = filterExpression();
            try
            {
                SqlDataSource.ConnectionString = strConnString();
                SqlDataSource.FilterExpression = filterExpression();
                SqlDataSource.DataBind();
                this.GridView1.DataBind();
                return;
            }
            catch(Exception ex)
            {
                Label5.Text = ex.Message + "" + ex.StackTrace;
            }


            

           // g_table.Columns.Add("UID");
            //g_table.Columns.Add("帳號");
            //g_table.Columns.Add("姓名");
            //g_table.Columns.Add("暱稱");
            //g_table.Columns.Add("電話");
            //g_table.Columns.Add("e-mail");
            //g_table.Columns.Add("總星星數", typeof(Int32));
            //g_table.Columns.Add("總分數",typeof(Int32));
            //g_table.Columns.Add("菜葉甘藷分數");
            //g_table.Columns.Add("孤挺花分數");
            //g_table.Columns.Add("毛豆分數");
            //g_table.Columns.Add("海芋分數");

            //using (SqlConnection connection = new SqlConnection(strConnString()))
            //{
            //    connection.Open();
            //    string SQL = "select * from LBG_GUEST g,[mGIPcoanew]..[Member] m where g.STATUS <> -1 and g.USERID = m.account";

            //    SqlCommand cmd = new SqlCommand();
            //    cmd.CommandText = SQL;
            //    cmd.Connection = connection;
            //    SqlDataReader result = cmd.ExecuteReader();
            //    if (!result.Read()) return ;
            //    try
            //    {
            //        int count = 0;
            //        do
            //        {

            //            g_table.Rows.Add( //result["UID"], 
            //                               result["USERID"], result["REALNAME"], result["NICKNAME"],
            //                              result["mobile"], result["email"], getSTAR(result["UID"].ToString()), getTOTALSCORE(result["UID"].ToString())
            //                              , getPLANTSCORE(result["UID"].ToString(), "A"), getPLANTSCORE(result["UID"].ToString(), "B"), getPLANTSCORE(result["UID"].ToString(), "C"), getPLANTSCORE(result["UID"].ToString(), "D"));
            //        }
            //        while (result.Read());
            //    }
            //    catch
            //    {
            //        return ;
            //    }
            //    finally
            //    {
            //        cmd.Cancel();
            //        result.Close();
            //        connection.Close(); 
            //    }
            //}
            //dataView = new DataView(g_table);

            //dataView.Sort = "總分數 DESC , 總星星數 DESC";
            ////this.GridView1.DataSource = dataView;
            //this.GridView1.DataBind();
        }

    

       
        public void SP_SCORE_SUMMARY()
        {
            using (SqlConnection connection = new SqlConnection(strConnString()))
            {
                connection.Open();
                //DB.SP_CLOSE_PLANT(SESSIONID);
                string SQL = "SP_SCORE_SUMMARY";
                //SqlConnection connection = this.openConn();
                if (connection == null) return;
                SqlCommand cmd = new SqlCommand();
                try
                {
                    
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandText = SQL;
                    cmd.Connection = connection;
                    SqlDataReader result = cmd.ExecuteReader();

                    result.Close();
                }
                catch
                {
                    return;
                }
                finally
                {
                    connection.Close();
                    //releaseConnection(connection);
                }
            }


            //DB.SP_CHECKOUT(SESSIONID);
        }
        protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            return;
        }

        protected void GridView1_PageIndexChanging(object sender, EventArgs e)
        {
        }

        protected void GridView1_PageIndexChanging1(object sender, GridViewPageEventArgs e)
        {
            this.GridView1.PageIndex = e.NewPageIndex;
            BindGrid();
            this.GridView1.DataBind();
        }

        protected void GridView1_Sorting(object sender, GridViewSortEventArgs e)
        {
           
        }

        protected void btn_Refresh_Click(object sender, EventArgs e)
        {
            BindGrid();
        }

        protected void btn_Excel_Click(object sender, EventArgs e)
        {
            
            BindGrid();
            this.ExportExcel();
        }

        protected void SqlDataSource_Selecting(object sender, SqlDataSourceSelectingEventArgs e)
        {

        }

        protected void btn_Reset_Click(object sender, EventArgs e)
        {
            this.txtName.Text = "";
            this.txtNickName.Text = "";
            this.txtScoreHighBound.Text = "";
            this.txtScoreLowBound.Text = "";
        }

        protected void btn_Renew_Refresh_Click(object sender, EventArgs e)
        {
            try
            {
                Label5.Text = "程序執行中....";
                this.SP_SCORE_SUMMARY();
            }
            finally
            {
                Label5.Text = "程序執行開始 , 預計於3 ~6 分鐘內完成...";
            }
        }
    }
}
