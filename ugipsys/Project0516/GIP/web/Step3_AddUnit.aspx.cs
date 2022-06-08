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
using GSS.Vitals.COA.Data;

public partial class GIP_web_Step3_AddUnit : System.Web.UI.Page
{
    public int CurrentRootId
    {
        get { return (int)Session["User_id"]; }
    }

    public int CurrentParentId
    {
        get { return (ViewState["CurrentParentId"] != null ? (int)ViewState["CurrentParentId"] : 0); }
        set { ViewState["CurrentParentId"] = value; }
    }

    public IList PersonalDataCategories
    {
        get
        {
            if (ViewState["PersonalDataCategories"] == null)
                ViewState["PersonalDataCategories"] = new ArrayList();
            return (IList)ViewState["PersonalDataCategories"];
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {

        if (!IsPostBack)
        {
            if (Request.QueryString["level"] != null)
            {

                Session.Add("datalevel", "1");
            }
            else
            {
                Session.Add("datalevel", "2");
            }
            CurrentParentId = Convert.ToInt32(Request.QueryString["parent"]);

            DataCategoryTypeDropDownList.DataSource = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getDataCategoryTypes();
            DataCategoryTypeDropDownList.DataBind();

            ListFieldsCheckBoxList.DataSource = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getAvailiableFields();
            ListFieldsCheckBoxList.DataBind();
            ListFieldsCheckBoxList.Items.Insert(0, new ListItem("標題", "xtitle", false));

            ContentFieldsCheckBoxList.DataSource = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getAvailiableFields();
            ContentFieldsCheckBoxList.DataBind();
            ContentFieldsCheckBoxList.Items.Insert(0, new ListItem("標題", "xtitle", false));

            NodeNameTextBox.Text = Request.QueryString["name"];
            NodeNameMemoTextBox.Text = Request.QueryString["memo"];
            IsNodeOpenRadioButtonList.SelectedValue = Request.QueryString["open"];
        }

        RenderCategoryMaintainPanel();
    }

    protected void UnitTypeDropDownList_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (UnitTypeDropDownList.SelectedValue != "")
        {
            ListTypePanel.Enabled = !(UnitTypeDropDownList.SelectedValue == "CP");
            ContentTypePanel.Enabled = true;
        }
        else
        {
            ListTypePanel.Enabled = false;
            ContentTypePanel.Enabled = false;
        }

        ListTypePanel.Visible = ListTypePanel.Enabled;
        ContentTypePanel.Visible = ContentTypePanel.Enabled;
    }

    protected void AdvanceConfigButton_Click(object sender, EventArgs e)
    {
        BasicConfigPanel.Visible = false;
        AdvanceConfigPanel.Visible = true;
    }


    protected void BasicConfigButton_Click(object sender, EventArgs e)
    {
        BasicConfigPanel.Visible = true;
        AdvanceConfigPanel.Visible = false;
    }

    protected void InsertButton_Click(object sender, EventArgs e)
    {
        IDictionary map = new Hashtable();

        map.Add("rootId", CurrentRootId);
        map.Add("parentId", CurrentParentId);
        map.Add("userId", Session["Name"].ToString());
        map.Add("nodeName", NodeNameTextBox.Text);
        map.Add("catNameMemo", NodeNameMemoTextBox.Text);
        map.Add("isOpen", (IsNodeOpenRadioButtonList.SelectedValue == "Y"));
        map.Add("nodeType", UnitTypeDropDownList.SelectedValue);
        map.Add("inUse", true);
        map.Add("Genname", Session["Genname"].ToString());
        map.Add("datalevel", Convert.ToInt32(Session["datalevel"].ToString()));
        Session.Remove("datalevel");
        string listStyle = "";
        if (ListPageLayoutRadioButton1.Checked)
        {
            listStyle = "style1";
        }
        else if (ListPageLayoutRadioButton2.Checked)
        {
            listStyle = "style2";
        }


        map.Add("listStyle", listStyle);
        map.Add("listFields", getCheckBoxListValues(ListFieldsCheckBoxList));

        //---vincent modify---2008/12/25---
        //---新增styleD,styleE,styleF---三種版型---

        string contentStyle = "";
        if (ContentPageLayoutRadioButton1.Checked)
        {
            contentStyle = "styleA";
        }
        else if (ContentPageLayoutRadioButton2.Checked)
        {
            contentStyle = "styleB";
        }
        else if (ContentPageLayoutRadioButton3.Checked)
        {
            contentStyle = "styleC";
        }
        else if (ContentPageLayoutRadioButton4.Checked)
        {
            contentStyle = "styleD";
        }
        else if (ContentPageLayoutRadioButton5.Checked)
        {
            contentStyle = "styleE";
        }
        else if (ContentPageLayoutRadioButton6.Checked)
        {
            contentStyle = "styleF";
        }
        else if (ContentPageLayoutRadioButton7.Checked)
        {
            contentStyle = "styleG";
        }
        //---end modify by vincent---

        map.Add("contentStyle", contentStyle);
        map.Add("contentFields", getCheckBoxListValues(ContentFieldsCheckBoxList));
        map.Add("dataCategoryType", DataCategoryTypeDropDownList.SelectedValue);
        map.Add("dataCategories", getListBoxValues(DataCategoriesListBox));

        try
        {
            CatelogTreeNode node = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().addCatelogNode(map);
            ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert('新增成功');location.href='Step3.aspx';", true);
        }
        catch (Exception ex)
        {
            throw new Exception("Error", ex);
            //ClientScript.RegisterClientScriptBlock(Page.GetType(), "Error", "alert(\"發生錯誤\\n" + ex.Message + "\");", true);
        }

    }

    protected void DataCategoryTypeDropDownList_SelectedIndexChanged(object sender, EventArgs e)
    {
        string selectedValue = DataCategoryTypeDropDownList.SelectedValue;

        IList list;

        if (selectedValue == "")
        {
            list = new ArrayList();
        }
        else if (selectedValue == "CUSTOM")
        {
            list = PersonalDataCategories;
        }
        else
        {
            list = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getDefaultDataCategories(selectedValue);
        }

        DataCategoriesListBox.DataSource = list;
        DataCategoriesListBox.DataBind();
    }

    protected void RenderCategoryMaintainPanel()
    {
        CategoryMaintainPanel.Visible = (DataCategoryTypeDropDownList.SelectedValue == "CUSTOM");
    }

    protected IList getCheckBoxListValues(CheckBoxList list)
    {
        IList results = new ArrayList();

        foreach (ListItem item in list.Items)
        {
            if (item.Selected)
                results.Add(item.Value);
        }

        return results;
    }

    protected IList getListBoxValues(ListBox list)
    {
        IList results = new ArrayList();

        foreach (ListItem item in list.Items)
        {
            results.Add(item.Value);
        }

        return results;
    }

    protected void AddCategoryButton_Click(object sender, EventArgs e)
    {
        // 新增個人化資料大類
        PersonalDataCategories.Add(new DictionaryEntry(CategoryNameTextBox.Text, CategoryNameTextBox.Text));

        DataCategoriesListBox.DataSource = PersonalDataCategories;
        DataCategoriesListBox.DataBind();
    }

    protected void RemoveCategoryButton_Click(object sender, EventArgs e)
    {
        // 刪除個人化資料大類
        foreach (ListItem item in DataCategoriesListBox.Items)
        {
            if (item.Selected)
            {
                foreach (DictionaryEntry entry in PersonalDataCategories)
                {
                    if ((string)entry.Key == item.Value)
                    {
                        PersonalDataCategories.Remove(entry);
                        break;
                    }
                }
            }
        }

        DataCategoriesListBox.DataSource = PersonalDataCategories;
        DataCategoriesListBox.DataBind();
    }

    protected void CategoryMoveUpButton_Click(object sender, EventArgs e)
    {
        // 上移個人化資料大類
        DictionaryEntry cuurentEntry;
        foreach (DictionaryEntry entry in PersonalDataCategories)
        {
            if ((string)entry.Key == DataCategoriesListBox.SelectedValue)
            {
                cuurentEntry = entry;
                break;
            }
        }

        int pos = PersonalDataCategories.IndexOf(cuurentEntry);

        if (pos == -1)
            return;

        if (pos > 0)
        {
            PersonalDataCategories.RemoveAt(pos);
            PersonalDataCategories.Insert(pos - 1, cuurentEntry);
        }

        DataCategoriesListBox.DataSource = PersonalDataCategories;
        DataCategoriesListBox.DataBind();
    }

    protected void CategoryMoveDownButton_Click(object sender, EventArgs e)
    {
        // 下移個人化資料大類
        DictionaryEntry cuurentEntry;
        foreach (DictionaryEntry entry in PersonalDataCategories)
        {
            if ((string)entry.Key == DataCategoriesListBox.SelectedValue)
            {
                cuurentEntry = entry;
                break;
            }
        }

        int pos = PersonalDataCategories.IndexOf(cuurentEntry);

        if (pos == -1)
            return;

        if (pos < PersonalDataCategories.Count - 1)
        {
            PersonalDataCategories.RemoveAt(pos);
            PersonalDataCategories.Insert(pos + 1, cuurentEntry);
        }

        DataCategoriesListBox.DataSource = PersonalDataCategories;
        DataCategoriesListBox.DataBind();
    }

    protected void AddCatelogFolderImageButton_Click(object sender, ImageClickEventArgs e)
    {
        Response.Redirect("Step3_AddFolder.aspx");
    }
    protected void cancel_Click(object sender, EventArgs e)
    {
        Response.Redirect("Step3.aspx");
    }
}
