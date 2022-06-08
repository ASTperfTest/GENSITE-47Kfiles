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
using Hyweb.M00.COA.GIP.TopicWeb;
using GSS.Vitals.COA.Data;

public partial class GIP_web_Step3_EditUnit : System.Web.UI.Page
{

	public CatelogNode CurrentNode
	{
		get { return (CatelogNode)ViewState["CurrentNode"]; }
		set { ViewState["CurrentNode"] = value; }
	}

	public IList PersonalDataCategories
	{
		get
		{
			if (ViewState["PersonalDataCategories"] == null)
				ViewState["PersonalDataCategories"] = new ArrayList();
			return (IList)ViewState["PersonalDataCategories"];
		}
		set
		{
			ViewState["PersonalDataCategories"] = value;
		}
	}

    //新聞擷取新增關鍵字
    public IList NewAddList
    {
        get
        {
            if (ViewState["NewAddList"] == null)
                ViewState["NewAddList"] = new ArrayList();
            return (IList)ViewState["NewAddList"];
        }
    }
    //新聞擷取排除關鍵字
    public IList NewORList
    {
        get
        {
            if (ViewState["NewORList"] == null)
                ViewState["NewORList"] = new ArrayList();
            return (IList)ViewState["NewORList"];
        }
    }

	protected void Page_Load(object sender, EventArgs e)
	{
		if (!IsPostBack)
		{
			// 載入單元資料
			int nodeId = Convert.ToInt32(Request.QueryString["nodeId"]);
			CurrentNode = TopicWebHelper.getInstance().getCatelogNode(nodeId);
            //2011/05/24 sam_chen :bob說要隱藏用不到的分類
            IList dataCategoryList = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getDataCategoryTypes();
            IList currentList = new ArrayList();
            foreach (DictionaryEntry dlist in dataCategoryList)
            {
                if (string.Compare(dlist.Key.ToString(), "CUSTOM", true) == 0)
                {
                    currentList.Add(dlist);
                    break;
                }
            }
            DataCategoryTypeDropDownList.DataSource = currentList;
			DataCategoryTypeDropDownList.DataBind();

			DataCategoryTypeDropDownList.SelectedValue = CurrentNode.DataCategoryType;

			// 列表欄位
			ListFieldsCheckBoxList.DataSource = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getAvailiableFields();
			ListFieldsCheckBoxList.DataBind();
			ListFieldsCheckBoxList.Items.Insert(0, new ListItem("標題", "xtitle", false));

			// 內容欄位
			ContentFieldsCheckBoxList.DataSource = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getAvailiableFields();
			ContentFieldsCheckBoxList.DataBind();
			ContentFieldsCheckBoxList.Items.Insert(0, new ListItem("標題", "xtitle", false));

			// 載入個人化資料大類
			PersonalDataCategories = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().getCustomDataCategories(CurrentNode.UnitId);

            //新聞擷取
            GetNewList();
			//DataCategoriesListBox.DataSource = PersonalDataCategories;
			//DataCategoriesListBox.DataBind();


			NodeNameTextBox.Text = CurrentNode.Name;
			NodeNameMemoTextBox.Text = CurrentNode.CatNameMemo;
			IsNodeOpenRadioButtonList.SelectedValue = CurrentNode.InUse ? "Y" : "N";
			UnitTypeDropDownList.SelectedValue = CurrentNode.UnitKind;
			CheckBoxListHelper.setSelected(ListFieldsCheckBoxList, CurrentNode.ListFields);
			ListPageLayoutRadioButton1.Checked = CurrentNode.ListStyle == "style1";
			ListPageLayoutRadioButton2.Checked = CurrentNode.ListStyle == "style2";
			
			CheckBoxListHelper.setSelected(ContentFieldsCheckBoxList, CurrentNode.ContentFields);
			DataCategoryTypeDropDownList.SelectedValue = CurrentNode.DataCategoryType;
			ContentPageLayoutRadioButton1.Checked = CurrentNode.ContentStyle == "styleA";
			ContentPageLayoutRadioButton2.Checked = CurrentNode.ContentStyle == "styleB";
			ContentPageLayoutRadioButton3.Checked = CurrentNode.ContentStyle == "styleC";
      //---vincent modify---2008/12/25---
      //---新增styleD,styleE,styleF---三種版型---
      ContentPageLayoutRadioButton4.Checked = CurrentNode.ContentStyle == "styleD";
      ContentPageLayoutRadioButton5.Checked = CurrentNode.ContentStyle == "styleE";
      ContentPageLayoutRadioButton6.Checked = CurrentNode.ContentStyle == "styleF";
	  
	  ContentPageLayoutRadioButton7.Checked = CurrentNode.ContentStyle == "styleG";
	  
      //---enf of modify by vincent---
			RenderUnitType();
			RenderDataCategory();
			RenderCategoryMaintainPanel();
		}
	}

	protected void RenderUnitType()
	{
		if (UnitTypeDropDownList.SelectedValue != "")
		{
			ListTypePanel.Visible = (UnitTypeDropDownList.SelectedValue == "LP");
			ContentTypePanel.Visible = true;
            //NewsPanel.Visible = (UnitTypeDropDownList.SelectedValue == "LP");
            if (UnitTypeDropDownList.SelectedValue == "LP")
            {
                NewsDiv.Visible = true;
                if (NewsDDL.SelectedValue == "Y")
                {
                    NewsPanel.Visible = true;
                }
                else
                {
                    NewsPanel.Visible = false;
                }
            }
            else
            {
                NewsDiv.Visible = false;
                NewsPanel.Visible = false;
            }
		}
		else
		{
			ListTypePanel.Visible = false;
			ContentTypePanel.Visible = false;
            NewsPanel.Visible = false;
		}
	}

	protected void RenderDataCategory()
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

	protected void UnitTypeDropDownList_SelectedIndexChanged(object sender, EventArgs e)
	{
		RenderUnitType();
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

	protected void DataCategoryTypeDropDownList_SelectedIndexChanged(object sender, EventArgs e)
	{
		RenderDataCategory();
		RenderCategoryMaintainPanel();
	}

    //2011/05/24 sam_chem 修正資料大類存檔時應該要分key跟value 
	protected IList getListBoxValues(ListBox list)
	{
		IList results = new ArrayList();

		foreach (ListItem item in list.Items)
		{
			results.Add(new DictionaryEntry(item.Value, item.Text));
		}

		return results;
	}


	protected void AddCategoryButton_Click(object sender, EventArgs e)
	{
		// 新增個人化資料大類
        //2011/05/24 sam_chen 新增資料大類時key用數字
        int oldCatId = 0;
        int newCatId = 0;
        foreach (DictionaryEntry entry in PersonalDataCategories)
        {
            if (int.TryParse((string)entry.Key, out oldCatId))
            {
                if (oldCatId > newCatId)
                    newCatId = oldCatId;
            }
        }
        newCatId++;
        string newCatkey = newCatId.ToString();
        if (newCatkey.Length == 1)
            newCatkey = "0" + newCatkey;
        //2011/05/24 sam_chen
        //避免長度超過2會造成抓不到文章 但是應該不可能修改100次吧???
        if (newCatkey.Length > 2)
            newCatkey = newCatkey.Substring(0, 2);
        PersonalDataCategories.Add(new DictionaryEntry(newCatkey, CategoryNameTextBox.Text));

		DataCategoriesListBox.DataSource = PersonalDataCategories;
		DataCategoriesListBox.DataBind();
	}

	protected void RemoveCategoryButton_Click(object sender, EventArgs e)
	{
		// 刪除個人化資料大類
       	foreach (ListItem item in DataCategoriesListBox.Items)
		{
			if (item.Selected == true)
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
		foreach(DictionaryEntry entry in PersonalDataCategories)
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

	protected void UpdateButton_Click(object sender, EventArgs e)
	{
		// 編修存檔
		IDictionary map = new Hashtable();
		map.Add("nodeId", CurrentNode.Id);
		map.Add("userId", "test");
		map.Add("nodeName", NodeNameTextBox.Text);
		map.Add("catNameMemo", NodeNameMemoTextBox.Text);
		map.Add("isOpen", (IsNodeOpenRadioButtonList.SelectedValue == "Y"));
		map.Add("nodeType", UnitTypeDropDownList.SelectedValue);
        map.Add("Genname", "test");

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
		map.Add("listFields", CheckBoxListHelper.getSelectedValues(ListFieldsCheckBoxList));

    //---vincent modify---2008/12/25---
    //---新增styleD,styleE,styleF---三種版型---

	//---Max modify---2011/05/17---
    //---新增styleG---一種版型---

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
		map.Add("contentFields", CheckBoxListHelper.getSelectedValues(ContentFieldsCheckBoxList));
		map.Add("dataCategoryType", DataCategoryTypeDropDownList.SelectedValue);
		map.Add("dataCategories", getListBoxValues(DataCategoriesListBox));

		try
		{
			CatelogTreeNode node = Hyweb.M00.COA.GIP.TopicWeb.TopicWebHelper.getInstance().updateCatelogNode(map);
            Update_KnowledgeToSubject();

			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert('修改成功');location.href='Step3.aspx';", true);
		}
		catch (Exception ex)
		{
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Error", "alert(\"發生錯誤: \\n" + ex.Message + "\");", true);
		}
	}

	protected void DeleteButton_Click(object sender, EventArgs e)
	{
		// 刪除
		try
		{
			TopicWebHelper.getInstance().deleteCatelogNode(CurrentNode.Id);
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Success", "alert('刪除成功');location.href='Step3.aspx';", true);
		}
		catch (Exception ex)
		{
			ClientScript.RegisterClientScriptBlock(Page.GetType(), "Error", "alert(\"發生錯誤:\\n" + ex.Message + "\");", true);
		}
	}
    protected void cancel_Click(object sender, EventArgs e)
    {
        Response.Redirect("Step3.aspx");
    }

    protected void AddInButton_Click(object sender, EventArgs e)
    {
        // 新增新聞擷取
        NewAddList.Add(new DictionaryEntry(addIntext.Text, addIntext.Text));

        NewsaddList.DataSource = NewAddList;
        NewsaddList.DataBind();
    }
    protected void AddOutButton_Click(object sender, EventArgs e)
    {
        // 新增新聞擷取排除
        NewORList.Add(new DictionaryEntry(addOuttext.Text, addOuttext.Text));

        NewsOrList.DataSource = NewORList;
        NewsOrList.DataBind();
    }

    protected void DelInButton_Click(object sender, EventArgs e)
    {
        if (NewsaddList.SelectedIndex > -1)
        {
            // 移除新聞擷取
            DictionaryEntry cuurentEntry = new DictionaryEntry();
            foreach (DictionaryEntry entry in NewAddList)
            {
                if ((string)entry.Key == NewsaddList.SelectedValue)
                {
                    cuurentEntry = entry;
                    break;
                }
            }
            int pos = NewAddList.IndexOf(cuurentEntry);
            if (pos <= NewAddList.Count - 1)
            {
                NewAddList.RemoveAt(pos);
            }

            NewsaddList.DataSource = NewAddList;
            NewsaddList.DataBind();
        }
    }

    protected void DelOutButton_Click(object sender, EventArgs e)
    {
        if (NewsOrList.SelectedIndex > -1)
        {
            // 新增新聞擷取
            DictionaryEntry cuurentEntry = new DictionaryEntry();
            foreach (DictionaryEntry entry in NewORList)
            {
                if ((string)entry.Key == NewsOrList.SelectedValue)
                {
                    cuurentEntry = entry;
                    break;
                }
            }
            int pos = NewORList.IndexOf(cuurentEntry);
            if (pos <= NewORList.Count - 1)
            {
                NewORList.RemoveAt(pos);
            }

            NewsOrList.DataSource = NewORList;
            NewsOrList.DataBind();
        }
    }
    protected void GetNewList()
    {
        string strnewadd, strnewnot, ss;
        strnewnot = null;
        string strExpress = null;
        String SQLSelectScript = @"SELECT  * FROM KnowledgeToSubject where CTNODEID=" + CurrentNode.Id;
        DataTable dt = SqlHelper.GetDataTable("LAMBDA_ConnectionString", SQLSelectScript);        

        if (dt.Rows.Count > 0)
        {
            strExpress = dt.Rows[0]["EXPRESSION"].ToString();
        }
        if (!string.IsNullOrEmpty(strExpress))
        {
            ss = strExpress.ToUpper();
            if (ss.IndexOf("AND") > -1)
            {
                ss = strExpress.ToUpper().Replace("AND", ",");
                strnewadd = ss.Split(',')[0].ToString();
                strnewnot = ss.Split(',')[1].ToString();
            }
            else
            {
                strnewadd = ss;
            }
            if (!string.IsNullOrEmpty(strnewadd))
            {
                strnewadd = strnewadd.Replace("(", "");
                strnewadd = strnewadd.Replace(")", "");
                strnewadd = strnewadd.Replace("OR", ",");

                string[] addlist = strnewadd.Split(',');

                for (int j = 0; j < addlist.Length; j++)
                {
                    NewAddList.Add(new DictionaryEntry(addlist[j], addlist[j].Trim()));
                }
                NewsaddList.DataSource = NewAddList;
                NewsaddList.DataBind();
                NewsDDL.SelectedValue = "Y";
                if (!string.IsNullOrEmpty(strnewnot))
                {
                    strnewnot = strnewnot.Replace("(", "");
                    strnewnot = strnewnot.Replace(")", "");
                    strnewnot = strnewnot.Replace("NOT", "");
                    strnewnot = strnewnot.Replace("OR", ",");
                    string[] orlist = strnewnot.Split(',');
                    for (int j = 0; j < orlist.Length; j++)
                    {
                        NewORList.Add(new DictionaryEntry(orlist[j], orlist[j].Trim()));
                    }
                    NewsOrList.DataSource = NewORList;
                    NewsOrList.DataBind();
                }
            }
            else
            {
                NewsDDL.SelectedValue = "N";
                NewsPanel.Visible = false;
            }
        }

    }
    protected void Update_KnowledgeToSubject()
    {
        int i = 0;
        if (NewsDDL.SelectedValue == "Y")
        {
            if (NewAddList.Count > 0)
            {
                string strExpress = null;
                foreach (DictionaryEntry entry in NewAddList)
                {
                    if (i == 0)
                    {
                        strExpress = "(" + (string)entry.Key.ToString().Trim();
                    }
                    else
                    {
                        strExpress += " OR " + (string)entry.Key.ToString().Trim();

                    }
                    i++;
                }
                strExpress += ")";

                if (NewORList.Count > 0)
                {
                    if (NewORList.Count > 1)
                    {
                        strExpress += " and (NOT(";
                    }
                    else
                    {
                        strExpress += " and (NOT ";
                    }
                    i = 0;
                    foreach (DictionaryEntry entry in NewORList)
                    {
                        if (i == 0)
                        {
                            strExpress += (string)entry.Key.ToString().Trim();
                        }
                        else
                        {
                            strExpress += " OR " + (string)entry.Key.ToString().Trim();

                        }
                        i++;
                    }
                    if (NewORList.Count > 1)
                    {
                        strExpress += "))";
                    }
                    else
                    {
                        strExpress += ")";
                    }
                }
                else
                {
                    strExpress = strExpress.Replace("(", "");
                    strExpress = strExpress.Replace(")", "");
                }

                try
                {
                    string SQLSelectScript = @"SELECT  * FROM KnowledgeToSubject where CTNODEID=" + CurrentNode.Id;
                    DataTable dt = SqlHelper.GetDataTable("LAMBDA_ConnectionString", SQLSelectScript);
                    if (dt.Rows.Count > 0)
                    {
                        String SQLUpdateScript = @"Update KnowledgeToSubject set EXPRESSION = '" + strExpress + "' where CTNODEID=" + CurrentNode.Id;
                        SqlHelper.ExecuteNonQuery("LAMBDA_ConnectionString", SQLUpdateScript);
                        SQLSelectScript = @"SELECT  count(*) FROM KnowledgeToSubject where CTNODEID=" + CurrentNode.Id;
                        int scount = (int)SqlHelper.ReturnScalar("COA_GIPConnectionString", SQLSelectScript);
                        if (scount < 1)
                        {
                            string SQLInsertScript = @"Insert into KnowledgeToSubject(subjectId,subjectName,baseDSDId,baseDSDName,ctUnitId,ctUnitName,ctNodeId,ctNodeName,status)" +
                             "Values( " + dt.Rows[0]["subjectId"] + ",'" + dt.Rows[0]["subjectName"] + "'," + dt.Rows[0]["baseDSDId"] + ",'" + dt.Rows[0]["baseDSDName"] +
                             "'," + dt.Rows[0]["ctUnitId"] + ",'" + dt.Rows[0]["ctUnitName"] + "'," + dt.Rows[0]["ctNodeId"] + ",'" + dt.Rows[0]["ctNodeName"] +
                             "','" + dt.Rows[0]["status"] + "')";
                            SqlHelper.ExecuteNonQuery("COA_GIPConnectionString", SQLInsertScript);
                        }
                    }
                    else
                    {
                        string SQLInsertScript = null;
                        SQLSelectScript = @"select CatTreeRoot.ctRootId as subjectId,CatTreeRoot.ctRootName as subjectName,7 as baseDSDId,'網頁' as baseDSDName "+
                            ", CtUnit.CtUnitId, CtUnit.CtUnitName,CatTreeNode.ctNodeId,CatTreeNode.catName,'Y' as status, null as iCUItem " +
                            "from CatTreeNode inner join CatTreeRoot on CatTreeRoot.ctRootId = CatTreeNode.ctRootId " +
                            "inner join CtUnit on CtUnit.ctUnitId = CatTreeNode.ctUnitId where CatTreeNode.ctNodeid = " + CurrentNode.Id;
                        DataTable dt2 = SqlHelper.GetDataTable("COA_GIPConnectionString", SQLSelectScript);

                        if (dt2.Rows.Count > 0)
                        {

                            SQLInsertScript = @"
                            set nocount on
                            Insert into KnowledgeToSubject(subjectId,subjectName,baseDSDId,baseDSDName,ctUnitId,ctUnitName,ctNodeId,ctNodeName,status)" +
                             "Values( " + dt2.Rows[0]["subjectId"] + ",'" + dt2.Rows[0]["subjectName"] + "'," + dt2.Rows[0]["baseDSDId"] + ",'" + dt2.Rows[0]["baseDSDName"] +
                             "'," + dt2.Rows[0]["ctUnitId"] + ",'" + dt2.Rows[0]["ctUnitName"] + "'," + dt2.Rows[0]["ctNodeId"] + ",'" + dt2.Rows[0]["catName"] +
                             "','Y')";
                            SQLInsertScript += System.Environment.NewLine;
                            SQLInsertScript += "select SCOPE_IDENTITY() as newId";

                            int sid = Convert.ToInt32(SqlHelper.ReturnScalar("COA_GIPConnectionString", SQLInsertScript));                         

                            SQLInsertScript = @"insert into KnowledgeToSubject (ID,SUBJECTID,SUBJECTNAME,BASEDSDID,BASEDSDNAME,CTUNITID,CTUNITNAME,CTNODEID,CTNODENAME,STATUS,EXPRESSION,LIMITROWS)" +
                               " Values(" + sid + "," + dt2.Rows[0]["subjectId"] + ",'" + dt2.Rows[0]["subjectName"] + "'," + dt2.Rows[0]["baseDSDId"] + ",'" + dt2.Rows[0]["baseDSDName"] +
                            "'," + dt2.Rows[0]["ctUnitId"] + ",'" + dt2.Rows[0]["ctUnitName"] + "'," + dt2.Rows[0]["ctNodeId"] + ",'" + dt2.Rows[0]["catName"] +
                            "','Y','" + strExpress + "',100)";
                            SqlHelper.ExecuteNonQuery("LAMBDA_ConnectionString", SQLInsertScript);
                        }
                    }
                }
                catch (Exception ex)
                {
                    ClientScript.RegisterClientScriptBlock(Page.GetType(), "Error", "alert(\"Update Expression Error:\\n" + ex.ToString() + "\");", true);
                }

            }
        }
        else
        {
            try
            {
                String SQLDeleteScript = null;
                string SQLSelectScript = @"SELECT  * FROM KnowledgeToSubject where CTNODEID=" + CurrentNode.Id;
                DataTable dt = SqlHelper.GetDataTable("LAMBDA_ConnectionString", SQLSelectScript);
                if (dt.Rows.Count > 0)
                {
                    SQLDeleteScript = @"Delete KnowledgeToSubject where CTNODEID=" + CurrentNode.Id;
                    SqlHelper.ExecuteNonQuery("LAMBDA_ConnectionString", SQLDeleteScript);

                }
                SQLSelectScript = @"SELECT  count(*) FROM KnowledgeToSubject where CTNODEID=" + CurrentNode.Id;
                int scount = (int)SqlHelper.ReturnScalar("COA_GIPConnectionString", SQLSelectScript);
                if (scount > 0)
                {
                    SQLDeleteScript = @"Delete KnowledgeToSubject where CTNODEID=" + CurrentNode.Id;
                    SqlHelper.ExecuteNonQuery("COA_GIPConnectionString", SQLDeleteScript);
                }
            }
            catch (Exception ex)
            {
                ClientScript.RegisterClientScriptBlock(Page.GetType(), "Error", "alert(\"Update Expression Error:\\n" + ex.Message + "\");", true);
            }
        }
    }
    protected void NewsDDL_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (NewsDDL.SelectedValue == "Y")
        {
            NewsPanel.Visible =true;
        }
        else
        {
            NewsPanel.Visible = false;
        }
    }

}
