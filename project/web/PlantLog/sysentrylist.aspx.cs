using System;
using System.Collections;
using System.Web.UI;
using System.Web.UI.WebControls;
using PlantLog.Core;
using PlantLog.Core.Domain;
using PlantLog.Core.Service;
using System.Data;
using System.Xml;
using System.Xml.Serialization;
using System.Text;
using Newtonsoft.Json;

public partial class sysentrylist : System.Web.UI.Page
{
    private IPlantLogService plantLogService;
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

    protected void Page_Load(object sender, EventArgs e)
    {
        if (fromBackEnd)
        {
            if (plantLogService == null)
            {
                plantLogService = Utility.ApplicationContext["PlantLogService"] as IPlantLogService;
            }

            Source = SetSource();
        }
        else
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "MyScript",
                "history.go(-1);", true);
        }

        if (!IsPostBack)
        {
            useManagementMasterPage.Value = fromBackEnd.ToString();
        }
    }

    protected override void OnPreInit(EventArgs e)
    {
        if (!IsPostBack)
        {
            if (Request.ServerVariables["HTTP_REFERER"] != null && (Request.ServerVariables["HTTP_REFERER"].ToLower().IndexOf(WebUtility.GetAppSetting("kmwebsysSite").ToString()) > -1))
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
        Entry e = plantLogService.GetEntry(entryId);
        e.IsApprove = true;
        plantLogService.UpdateEntry(e);
        Source = SetSource();
        BindData();
    }

    private void HideEntry(string entryId)
    {
        Entry e = plantLogService.GetEntry(entryId);
        e.IsApprove = false;
        plantLogService.UpdateEntry(e);
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
            entryControl.IsEditable = false;
            entryControl.IsAdmin = true;
            entryControl.IsVote = true;
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
        IList result = plantLogService.GetAllEntry();

        DataTable dtTemp = new DataTable();
        dtTemp.Columns.Add(new DataColumn("OWNER_ID"));
        dtTemp.Columns.Add(new DataColumn("TITLE"));
        dtTemp.Columns.Add(new DataColumn("LastModifyDateTime", typeof(DateTime)));
        dtTemp.Columns.Add(new DataColumn("Entry"));

        foreach (Entry temp in result)
        {
            DataRow dr = dtTemp.NewRow();
            dr["OWNER_ID"] = temp.OwnerId;
            dr["TITLE"] = temp.Title;
            dr["LastModifyDateTime"] = temp.ModifyDateTime;
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
