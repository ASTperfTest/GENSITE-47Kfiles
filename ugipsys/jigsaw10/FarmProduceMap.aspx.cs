using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Transactions;
using System.Data.Linq;
using System.Data.SqlClient;
using System.Data;
using System.Text;
using GSS.Vitals.COA.Data;


public partial class FarmProduceMap : System.Web.UI.Page
{
    private PagedDataSource Pager;
    private DataTable dt;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            QueryAll();
            BindData(0, 15);
        }
    }

    private void QueryAll() 
    {
        Query();
        Session["FROM"] = "";
    }

    protected void FillUnitDropDowmList(object sender, EventArgs e)
    {
        string type = type_DropDownList.SelectedValue;
        if (type == "0" || type == "1")
        {
            StringBuilder sql = new StringBuilder();
            sql.Append("select iCUItem, sTitle from CuDTGeneric where iCTUnit='2199' \n");
            if (type == "0") // 作物
                sql.Append("and RSS='0' \n");
            else if (type == "1") // 魚種
                sql.Append("and RSS='1' \n");
            sql.Append("order by sTitle DESC");
            try
            {
                DataTable dt = SqlHelper.GetDataTable("mGIPcoanewConnectionString", sql.ToString());
                int count = dt.Rows.Count;
                unit_DropDownList.Items.Clear();
                unit_DropDownList.Items.Add(new ListItem("請選擇", ""));
                for (int i = 0; i < count; i++)
                {
                    unit_DropDownList.Items.Add(new ListItem(dt.Rows[i].ItemArray.GetValue(1).ToString(), dt.Rows[i].ItemArray.GetValue(0).ToString()));
                }
            }
            catch (Exception ex) { }
        }
        else
        {
            unit_DropDownList.Items.Clear();
            unit_DropDownList.Items.Add(new ListItem("請選擇", ""));
        }
        // 清空rpList
        rpList.DataSource = null;
        rpList.DataBind();
    }

    protected void UnitQuery() 
    {
        if (unit_DropDownList.SelectedValue != "")
        {
            string sql = "";
            sql = @"select B.sTitle, B.unitId, CuDTGeneric.icuitem, CuDTGeneric.iEditor, CuDTGeneric.Created_Date, CuDTGeneric.xBody, CuDTGeneric.fCTUPublic from (
                        select A.sTitle, A.icuitem as unitId, CuDTGeneric.icuitem as icuitem_1 from (
                          select CuDTGeneric.icuitem, KnowledgeJigsaw.gicuitem as gicuitem_1, CuDTGeneric.sTitle from CuDTGeneric 
                          left join KnowledgeJigsaw on CuDTGeneric.icuitem=KnowledgeJigsaw.parenticuitem
                          where icuitem='" + unit_DropDownList.SelectedValue + @"' 
                        ) as A 
                        left join CuDTGeneric on CuDTGeneric.icuitem = A.gicuitem_1
                        where topCat='F'
                        ) as B
                      left join KnowledgeJigsaw on KnowledgeJigsaw.parenticuitem=B.icuitem_1
                      left join CuDTGeneric on KnowledgeJigsaw.gicuitem = CuDTGeneric.icuitem
                      order by Created_Date DESC";
            try
            {
                dt = SqlHelper.GetDataTable("mGIPcoanewConnectionString", sql);
                if (dt.Rows.Count == 1 && dt.Rows[0][2].ToString() == "")
                {
                    dt = null;
                }
                Session["DATA"] = dt;
                Session["FROM"] = "ChangeUnit";
            }
            catch (Exception ex) { }
        }
    }
    protected void ChangeUnit(object sender, EventArgs e) 
    {
        UnitQuery();
        BindData(0, 15);
        unit.Text = "";
        content.Text = "";
        author.Text = "";
    }

    protected void Query()
    {
        string sql = "";
        sql = @"select B.sTitle, B.unitId, CuDTGeneric.icuitem, CuDTGeneric.iEditor, CuDTGeneric.Created_Date, CuDTGeneric.xBody, CuDTGeneric.fCTUPublic from (
                   select A.sTitle, A.unitId, icuitem as icuitem_1 from (
                     select CuDTGeneric.icuitem as unitId, KnowledgeJigsaw.gicuitem as gicuitem_1, CuDTGeneric.sTitle from CuDTGeneric 
                     left join KnowledgeJigsaw on CuDTGeneric.icuitem=KnowledgeJigsaw.parenticuitem";
        if (unit.Text.Trim() != "")
            sql += @"
                     where sTitle like '%" + unit.Text.Trim() + "%' ";
         sql += @"
                   ) as A 
                   left join CuDTGeneric on CuDTGeneric.icuitem = A.gicuitem_1
                   where topCat='F'
                   ) as B
                 left join KnowledgeJigsaw on KnowledgeJigsaw.parenticuitem=B.icuitem_1
                 left join CuDTGeneric on KnowledgeJigsaw.gicuitem = CuDTGeneric.icuitem
                 where CuDTGeneric.icuitem is not null ";
         if (author.Text.Trim() != "")
             sql += @"
                 and CuDTGeneric.iEditor like '%" + author.Text.Trim() + "%' ";
        if (content.Text.Trim() != "")
            sql += @"
                 and CuDTGeneric.xBody like '%" + content.Text.Trim() + "%' ";
        sql += @"
                 order by Created_Date DESC";
        
        try
        {
            dt = SqlHelper.GetDataTable("mGIPcoanewConnectionString", sql);
            Session["DATA"] = dt;
            Session["FROM"] = "Query_btn_Click";
        }
        catch (Exception ex) { }
    }

    protected void Query_btn_Click(object sender, EventArgs e)
    {
        Query();
        BindData(0, 15);
        type_DropDownList.SelectedValue = "";
        unit_DropDownList.Items.Clear();
        unit_DropDownList.Items.Add(new ListItem("請選擇", ""));
    }

    protected void ChangePage(object sender, EventArgs e)
    {
        BindData(Convert.ToInt32(nowpage.SelectedValue), Convert.ToInt32(pagesize.SelectedValue));
    }

    protected void ChangePageSize(object sender, EventArgs e)
    {
        BindData(0, Convert.ToInt32(pagesize.SelectedValue));
    }
    protected void preLinkAct(object sender, EventArgs e)
    {
        nowpage.SelectedIndex--;
        BindData(Convert.ToInt32(nowpage.SelectedValue), Convert.ToInt32(pagesize.SelectedValue));
    }

    protected void nextLinkAct(object sender, EventArgs e)
    {
        nowpage.SelectedIndex++;
        BindData(Convert.ToInt32(nowpage.SelectedValue), Convert.ToInt32(pagesize.SelectedValue));
    }


    protected void BindData(int pageNumber, int pageSize)
    {
        Session["PAGENUMBER"] = pageNumber;
        Session["PAGESIZE"] = pageSize;
        string from = (string)Session["FROM"];
        dt = (DataTable)Session["DATA"];
        if (dt == null && from == "ChangeUnit") 
        {
            rpList.DataSource = dt;
            rpList.DataBind();
            nowpage.Items.Clear();
            pageNum.Text = "0";
            totalPage.Text = "0";
            totalRecord.Text = "0";
        }
        else if (dt == null)
        {
            Response.Write("<script>alert('請重新查詢！')</script>");
        }
        else 
        {
            Pager = dt.Paging(pageNumber, pageSize);
            rpList.DataSource = Pager;
            rpList.DataBind();
            SetControl();
        }
        
    }

    protected void SetControl()
    {
        nowpage.Items.Clear();
        for (int i = 0; i < Pager.PageCount; i++)
        {
            nowpage.Items.Add(new ListItem((i + 1).ToString(), i.ToString()));
        }
        nowpage.SelectedIndex = Pager.CurrentPageIndex;

        if (nowpage.Items.Count > 1)
        {
            if (Pager.IsFirstPage)
            {
                preLink.Visible = false;
                nextLink.Visible = true;
            }
            else if (Pager.IsLastPage)
            {
                preLink.Visible = true;
                nextLink.Visible = false;
            }
            else
            {
                preLink.Visible = true;
                nextLink.Visible = true;
            }
        }
        else
        {
            preLink.Visible = false;
            nextLink.Visible = false;
        }
        pageNum.Text = (Pager.CurrentPageIndex + 1).ToString();
        totalPage.Text = Pager.PageCount.ToString();
        totalRecord.Text = dt.Rows.Count.ToString();
    }

    protected void Edit(object sender, EventArgs e)
    {
        Button trigBtn = (Button)sender;
        foreach (RepeaterItem item in this.rpList.Items)
        {
            Button editBtn = item.FindControl("Button1") as Button;
            Label lbl = item.FindControl("Label1") as Label;
            TextBox txt = item.FindControl("TextBox1") as TextBox;
            Button saveBtn = item.FindControl("save") as Button;
            if (editBtn.ClientID == trigBtn.ClientID)
            {
                lbl.Style["display"] = "none";
                txt.Style["display"] = "inline";
                saveBtn.Style["display"] = "inline";
                txt.Text = lbl.Text;
            }
        }
    }

    protected void Save(object sender, EventArgs e)
    {
        Button trigBtn = (Button)sender;
        foreach (RepeaterItem item in this.rpList.Items)
        {
            Label lbl = item.FindControl("Label1") as Label;
            TextBox txt = item.FindControl("TextBox1") as TextBox;
            Button saveBtn = item.FindControl("save") as Button;
            if (saveBtn.ClientID == trigBtn.ClientID)
            {
                txt.Style["display"] = "none";
                saveBtn.Style["display"] = "none";
                lbl.Style["display"] = "inline";
                lbl.Text = txt.Text;
                string sql = "update CuDTGeneric set xBody='" + txt.Text + "' where iCUItem='" + saveBtn.ToolTip + "'";
                SqlHelper.ReturnScalar("mGIPcoanewConnectionString", sql);

            }
        }
    }

    protected string FixYN(object str)
    {
        if (str.ToString() == "Y")
            return "公開";
        else if (str.ToString() == "N")
            return "不公開";
        else
            return "";
    }

    protected void DealPublic(object sender, EventArgs e)
    {
        bool selected = false;
        Button btn = (Button)sender;
        string fCTUPublic = "";
        if (btn.ID == "public")
            fCTUPublic = "Y";
        else
            fCTUPublic = "N";

        foreach (RepeaterItem item in this.rpList.Items)
        {
            CheckBox ckb = item.FindControl("CheckBox1") as CheckBox;
            if (ckb.Checked == true)
                selected = true;
        }

        if (!selected)
        {
            Response.Write("<script>alert('請至少選擇一項！');</script>");
        }
        else
        {
            try
            {
                foreach (RepeaterItem item in this.rpList.Items)
                {
                    CheckBox ckb = item.FindControl("CheckBox1") as CheckBox;
                    if (ckb.Checked == true)
                    {
                        string sql = "update CuDTGeneric set fCTUPublic='" + fCTUPublic + "' where iCUItem='" + ckb.ToolTip + "'";
                        SqlHelper.ReturnScalar("mGIPcoanewConnectionString", sql);
                    }
                }
                if (Session["FROM"] == "Query_btn_Click")
                    Query();
                else if (Session["FROM"] == "ChangeUnit")
                    UnitQuery();
                else 
                    QueryAll();
                BindData((Int32)Session["PAGENUMBER"], (Int32)Session["PAGESIZE"]);
            }
            catch (Exception ex) { }
        }
    }
}

