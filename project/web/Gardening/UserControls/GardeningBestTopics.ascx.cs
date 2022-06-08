using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Collections;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;
using System.Data;
using System.Xml;
using System.Xml.Serialization;


public partial class GardeningBestTopics : System.Web.UI.UserControl
{
    private IGardeningService gardeningService;
    //預設取前5筆
    private int order=5;
    public int Order
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
        if (gardeningService == null)
        {
            gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        }

        BestTopicsDataList.DataSource = GetSource();
        BestTopicsDataList.DataBind();
    }

    private DataTable GetSource()
    {
        IList result = gardeningService.GetTopicsByRecommendedOrder(this.order);

        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add(new DataColumn("ImageUri"));
		dtTemp.Columns.Add(new DataColumn("TopicUri"));
        dtTemp.Columns.Add(new DataColumn("Title"));
        dtTemp.Columns.Add(new DataColumn("Description"));

        if (result.Count != 0)
        {	
			foreach( Topic temp in result)
			{
				DataRow dr = dtTemp.NewRow();
				dr["ImageUri"] = @"~\Gardening\" + temp.Avatar.Uri + @"\" + temp.Avatar.Name;
				dr["TopicUri"] = @"~\Gardening\entrylist.aspx?topicid=" + temp.TopicId;
				dr["Title"] = temp.Title;
				
				if (temp.Title.Length > 8)
					dr["Title"] = temp.Title.Substring(0, 8) + "....";
				else
					dr["Title"] = temp.Title;				
				
				if (temp.Description.Length > 20)
					dr["Description"] = temp.Description.Substring(0, 20) + "....";
				else
					dr["Description"] = temp.Description;						

				dtTemp.Rows.Add(dr);
			}			
        }
        else
        {
            DataRow dr = dtTemp.NewRow();
            dr["Title"] = "目前無資料";
            dtTemp.Rows.Add(dr);
        }
        return dtTemp;
    }
}
