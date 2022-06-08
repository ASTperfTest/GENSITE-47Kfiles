using System;
using System.Data;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

public partial class GardeningExpert : System.Web.UI.UserControl
{
    
	//預設取ORDER 3 之前的達人
    public string order="3";
    public string Order
    {
        get
        { 
            return this.order; 
        }
        set
        { 
            this.order = value; 
        }
    }
	
	protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            DataTable dt = GetGardenExpert();
            gardenExpertDataList.DataSource = dt;
            gardenExpertDataList.DataBind();
        }
    }

    //取出園藝達人前3筆資料
    private DataTable GetGardenExpert()
    {
        //先做出空data table和col，之後再一筆一筆加row
        DataTable dt = new DataTable();
        dt.Columns.Add(new DataColumn("Image"));
        dt.Columns.Add(new DataColumn("Name"));
        dt.Columns.Add(new DataColumn("Intro"));

        //sql抓資料
        string connString = System.Web.Configuration.WebConfigurationManager.AppSettings["COAConnectionString"];
        string Command = "SELECT TOP  " + Order + @" INTRODUCTION, realname, photo FROM GARDENING_EXPERT AS A JOIN Member AS B
	                            ON A.ACCOUNT = B.ACCOUNT	
								ORDER BY SORT_ORDER";

        using (SqlConnection sqlConnObj = new SqlConnection(connString))
        {
            //設定查詢語法
            sqlConnObj.Open();
            SqlCommand tCmd = new SqlCommand(Command, sqlConnObj);

            //抓回資料
            SqlDataAdapter da = new SqlDataAdapter(tCmd);
            DataTable gardenExpertDt = new DataTable();
            da.Fill(gardenExpertDt);

            //園藝達人資料列
            if (gardenExpertDt.Rows.Count != 0)
            {
				for (int i = 0; i < gardenExpertDt.Rows.Count; i++)
			    {
                    DataRow dr = dt.NewRow();
					if (Convert.ToString(gardenExpertDt.Rows[i]["photo"])=="")
                        dr["Image"] = "~/Gardening/images/default.jpg";
                    else
                        dr["Image"] = "~" + Convert.ToString(gardenExpertDt.Rows[i]["photo"]);
                    if (Convert.ToString(gardenExpertDt.Rows[i]["realname"])=="")
                        dr["Name"] = "[ 園藝王 ]";
                    else
                        dr["Name"] = "[ " + Convert.ToString(gardenExpertDt.Rows[i]["realname"]) + " ]";
					if ((Convert.ToString(gardenExpertDt.Rows[i]["INTRODUCTION"])).Length > 18)
                        dr["Intro"] = (Convert.ToString(gardenExpertDt.Rows[i]["INTRODUCTION"])).Substring(0, 18) + "....";
                    else
                        dr["Intro"] = Convert.ToString(gardenExpertDt.Rows[i]["INTRODUCTION"]);
					dt.Rows.Add(dr);
			    }                   
            }
			else
            { 
                    DataRow dr = dt.NewRow();
                    dr["Name"] = "目前無資料";
                    dt.Rows.Add(dr);
            }
        }             
        return dt;
    }
}
