using System;
using System.Collections;
using System.Web.UI;
using System.Web.UI.WebControls;
using Gardening.Core;
using Gardening.Core.Domain;
using Gardening.Core.Service;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.Text;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Data;

public partial class sysentrylist : System.Web.UI.Page
{
    private IGardeningService gardeningService;
	private bool fromBackEnd = false;
    
    private DataTable Source
    {
        get
        {
            if (ViewState["AllEntry"] == null)
            {
                return null;
            }
            else
            {
                return ViewState["AllEntry"] as DataTable;
            }
        }
        set
        {
            ViewState["AllEntry"] = value;
        }
    }
	
	protected override void OnPreInit(EventArgs e)
    {
        if (!IsPostBack)
        {
            if ((Request.ServerVariables["HTTP_REFERER"] != null && (Request.ServerVariables["HTTP_REFERER"].ToLower().IndexOf(WebUtility.GetAppSetting("kmwebsysSite").ToString()) > -1)) || (Session["FromBackEnd"] != null && Convert.ToBoolean(Session["FromBackEnd"])))
            {
                WebUtility.SetManagerSessions();
                fromBackEnd = true;
            }
        }
        else
        {
            string hiddenFromBackEnd = Request.Form["ctl00$cp$useManagementMasterPage"];
            if (hiddenFromBackEnd != null && hiddenFromBackEnd != "" && Convert.ToBoolean(hiddenFromBackEnd))
            {
                fromBackEnd = true;
            }
        }

        if (fromBackEnd)
        {
            this.MasterPageFile = WebUtility.managementMasterPageFile;
			Session["FromBackEnd"]=true.ToString();
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
		if (gardeningService == null)
		{
			gardeningService = Utility.ApplicationContext["GardeningService"] as IGardeningService;
		}

		Source = SetSource();
		
        if (!IsPostBack)
        {
            useManagementMasterPage.Value = fromBackEnd.ToString();
        }
    }

    protected void Page_PreRender(object sender, EventArgs e)
    {
        if (Request.Form["Action"] != null)
        {
            string act = Request.Form["Action"];
            if (act.Contains("ApproveEntry"))
            {
                ApproveEntry(act.Split('|')[1]);
            }

            if (act.Contains("HideEntry"))
            {
                HideEntry(act.Split('|')[1]);
            }
			if (act.Contains("DeleteEntry"))
			{
                DeleteEntry(act.Split('|')[1]);
            }

            if (act.Contains("EditEntry"))
            {
                EditEntry(act.Split('|')[1]);
            }
        }

        BindData();
    }

    private void BindData()
    {
        DataView dv = Source.DefaultView;
		
        dv.Sort = rdobtnSortBy.SelectedValue + " " + rdobtnOrder.SelectedValue;
        GridView1.DataSource = dv;
        GridView1.DataBind();

        SetGridBottomPagerInfo();
    }

    private void ApproveEntry(string entryId)
    {
        Entry e = gardeningService.GetEntry(entryId);
        e.IsApprove = true;
        gardeningService.UpdateEntry(e);
        Source = SetSource();
        BindData();
    }

    private void HideEntry(string entryId)
    {
        Entry e = gardeningService.GetEntry(entryId);
        e.IsApprove = false;
        gardeningService.UpdateEntry(e);
        Source = SetSource();
        BindData();
    }
	
	private void EditEntry(string entryId)
    {
		Entry thisEntry = gardeningService.GetEntry(entryId);
        Response.Redirect("singleentry.aspx?ownerid=" + thisEntry.CreatorId + "&entryid=" + entryId + "&topicid=" + thisEntry.TopicId);
    }

    private void DeleteEntry(string entryId)
    {
		//delete kpi data
		Entry thisEntry = gardeningService.GetEntry(entryId);
        string strCreateDate = thisEntry.CreateDateTime.ToShortDateString();
		//Response.Write("<script language=javascript>alert('" + strCreateDate +"!')</script>");
		
		using (SqlConnection tSqlConn = new SqlConnection())
        {
            string webConfigConnectionString1 = System.Web.Configuration.WebConfigurationManager.ConnectionStrings["ODBCDSN"].ConnectionString;

            tSqlConn.ConnectionString = webConfigConnectionString1;
            tSqlConn.Open();

            string sql = @"SELECT memberId FROM MemberGradeShare "+  
                                            "WHERE (memberId =@memberId)  " +
                                            "AND (CONVERT(nvarchar, shareDate, 111) = @creatDate)";
            SqlCommand Cmd = new SqlCommand(sql, tSqlConn);
            Cmd.Parameters.AddWithValue("memberId", Session["memID"]);
            Cmd.Parameters.AddWithValue("creatDate", strCreateDate);

            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommandBuilder cb = new SqlCommandBuilder();
            DataTable dt = new DataTable();

            da.SelectCommand = Cmd;
            cb.DataAdapter = da; // 指定DataAdapter
            da.Fill(dt);

            if ( dt != null && dt.Rows.Count >= 1)
            {
				Cmd.CommandText= @"update MemberGradeShare " +
                                "set sharegardening = case when sharegardening >=1 then sharegardening-1 else 0 end " +
                                "WHERE memberId=@memberId and (CONVERT(nvarchar, shareDate, 111) = @creatDate)";	
                
                Cmd.ExecuteNonQuery();
			}
            else
            { 
                // exception
            }
        }       
				
		gardeningService.DeleteEntry(entryId);		
        Source = SetSource();
        BindData();
    }

    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            UserControls_EntryControl entryControl = e.Row.FindControl("EntryControl1") as UserControls_EntryControl;
            string entryString = ((DataRowView)e.Row.DataItem)["Entry"].ToString();
            entryControl.Source = XmlDeserialize(entryString);
            entryControl.IsEditable = true;
            entryControl.IsAdmin = true;
            entryControl.IsOwner = false;
        }
    }

    protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        if (e.NewPageIndex >= 0)
        {
            GridView1.PageIndex = e.NewPageIndex;
            BindData();
        }
    }

    private DataTable SetSource()
    {
        IList result = gardeningService.GetAllEntries();

        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add(new DataColumn("OWNER_ID"));
        dtTemp.Columns.Add(new DataColumn("TITLE"));
        dtTemp.Columns.Add(new DataColumn("LastModifyDateTime", typeof(DateTime)));
        dtTemp.Columns.Add(new DataColumn("Entry"));
		dtTemp.Columns.Add(new DataColumn("Date", typeof(DateTime)));
        dtTemp.Columns.Add(new DataColumn("CreateDateTime", typeof(DateTime)));
		
        foreach (Entry temp in result)
        {
            DataRow dr = dtTemp.NewRow();
            dr["OWNER_ID"] = temp.CreatorId;
            dr["TITLE"] = temp.Title;
            dr["LastModifyDateTime"] = temp.ModifyDateTime;
            dr["Date"] = temp.Date;
            dr["CreateDateTime"] = temp.CreateDateTime;			
            dr["Entry"] = XmlSerialize(temp);
            dtTemp.Rows.Add(dr);
        }

        return dtTemp;
    }

    private string XmlSerialize(Entry entry)
    {
        XmlSerializer ser = new XmlSerializer(entry.GetType());
        StringBuilder sb = new StringBuilder();

        System.IO.StringWriter writer = new System.IO.StringWriter(sb);
        ser.Serialize(writer, entry);


        return sb.ToString();
    }
    private Entry XmlDeserialize(string s)
    {
        Entry entry = new Entry();
        XmlDocument xdoc = new XmlDocument();
        xdoc.LoadXml(s);
        XmlNodeReader reader = new XmlNodeReader(xdoc.DocumentElement);
        XmlSerializer ser = new XmlSerializer(entry.GetType());
        object obj = ser.Deserialize(reader);

        return obj as Entry;
    }

    protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
    {
        DropDownList numberOfTuplesDDL =
            GridView1.BottomPagerRow.Cells[0].FindControl("NumOfTuplesDropDownList") as DropDownList;
        GridView1.PageSize = Convert.ToInt32(numberOfTuplesDDL.SelectedValue);
        int selectedIndex = numberOfTuplesDDL.SelectedIndex;
        GridView1.PageIndex = 0;
        //BindGridView
        BindData();
    }

    protected void Sort_Click(object sender, EventArgs e)
    {
        Source = SetSource();
        BindData();
    }

    private void SetGridBottomPagerInfo()
    {
        Label pageTextBox = this.GridView1.BottomPagerRow.Cells[0].FindControl("PageIndexTextBox") as Label;
        pageTextBox.Text = Convert.ToString(this.GridView1.PageIndex + 1);

        Label totalTuplesTextBox = this.GridView1.BottomPagerRow.Cells[0].FindControl("TotalTuplesTextBox") as Label;
        totalTuplesTextBox.Text = Convert.ToString(Source.Rows.Count);

        Label totalPagesTextBox = this.GridView1.BottomPagerRow.Cells[0].FindControl("TotalPagesTextBox") as Label;
        totalPagesTextBox.Text = Convert.ToString(this.GridView1.PageCount);

        DropDownList numberOfTuplesDDL =
            this.GridView1.BottomPagerRow.Cells[0].FindControl("NumOfTuplesDropDownList") as DropDownList;

        switch (this.GridView1.PageSize)
        {
            case 5:
                numberOfTuplesDDL.SelectedIndex = 0;
                break;
            case 10:
                numberOfTuplesDDL.SelectedIndex = 1;
                break;
            case 50:
                numberOfTuplesDDL.SelectedIndex = 2;
                break;
            default:
                goto case 5;
        }
    }
}
