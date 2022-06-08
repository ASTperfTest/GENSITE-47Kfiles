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

public partial class GardeningDiary : System.Web.UI.UserControl
{

    private IGardeningService gardeningService;
    
    protected void Page_Load(object sender, EventArgs e)
    {
        if (gardeningService == null)
        {
            gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
        }

        DiaryDataList.DataSource = GetSource();
        DiaryDataList.DataBind();
    }

    private DataTable GetSource()
    {
        IList result = gardeningService.GetAllEntries();

        DataTable dtTemp = new DataTable();
		dtTemp.Columns.Add(new DataColumn("ImageUri"));
        dtTemp.Columns.Add(new DataColumn("TITLE"));
        dtTemp.Columns.Add(new DataColumn("LastModifyDateTime"));
		dtTemp.Columns.Add(new DataColumn("LastModifyDateTimeSort", typeof(DateTime)));

        //日誌
		if (result.Count != 0)
        {
            if(result.Count >= 6 )
			{
				for (int i = 0; i < 6; i++ )
				{
					Entry temp = (Entry)result[i];
					DataRow dr = dtTemp.NewRow();
					dr["ImageUri"] = "http://kminter.coa.gov.tw/gardening/entrylist.aspx?topicid=" + temp.TopicId;
					dr["TITLE"] = temp.Title;
					dr["LastModifyDateTime"] = temp.ModifyDateTime.ToShortDateString();
					dr["LastModifyDateTimeSort"] = temp.ModifyDateTime;
					dtTemp.Rows.Add(dr);
				}
			}
			else
			{
				foreach( Entry temp in result)
				{
					DataRow dr = dtTemp.NewRow();
					dr["ImageUri"] = "http://kminter.coa.gov.tw/gardening/entrylist.aspx?topicid=" + temp.TopicId;
					dr["TITLE"] = temp.Title;
					dr["LastModifyDateTime"] = temp.ModifyDateTime.ToShortDateString();
					dr["LastModifyDateTimeSort"] = temp.ModifyDateTime;
					dtTemp.Rows.Add(dr);
				}
			}			
            //依日誌更新最新時間排序
            DataView dv = dtTemp.DefaultView;
            dv.Sort = "LastModifyDateTimeSort DESC";
        }
        else
        {
            DataRow dr = dtTemp.NewRow();
            dr["TITLE"] = "目前無資料";
            dtTemp.Rows.Add(dr);
        }

        return dtTemp;
    }
}
