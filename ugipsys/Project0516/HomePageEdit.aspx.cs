using System;
using System.Data;
using System.Xml;
using System.Web.UI.WebControls;
using System.Web.UI;
using System.IO;

public partial class HomePageEdit : System.Web.UI.Page
{
    public CreateXdmpXml XmlFile = new CreateXdmpXml();   
   
    protected void Page_Load(object sender, EventArgs e)
    {
  Session["TreeRootId"] = Session["User_id"].ToString();
        string[] FileSubName = { "jpg", "jpeg", "gif", "png", "bmp" };
        for (int i = 0; i < FileSubName.Length; i++)
        {
            HomePageSql HomePageSql = new HomePageSql();
            string PicPath = "~/images/ugipnull/big/" + HomePageSql.LayoutPicName(Session["TreeRootId"].ToString()) + "." + FileSubName[i].ToString();
            if (File.Exists(Page.Server.MapPath(PicPath)))
            {
                LayoutPic.ImageUrl = PicPath;
                break;
            }
        }


      
        string FilePath = savePath();
        if (FilePath == "-1")
        {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "", "alert('找不到檔案')",true);
            return;                   
        }
        XmlAllData();        
        
    }

    //取得xml資料
   
    //顯示資料
    protected void XmlAllData()
    {
        string FilePath = savePath();  
        DataSet dsFormXml = XmlFile.loadXmlFile(FilePath);    
        HomePageSql nodeItem_Sql = new HomePageSql();
        DataTable nodeItem_dt = nodeItem_Sql.NodeItem(Convert.ToInt32(Session["TreeRootId"]));

        int n = dsFormXml.Tables[1].Rows.Count;

        UserControl BlockControls;
        for (int i = 0; i < n; i++)
        { 
            BlockControls = (UserControl)Page.LoadControl("BlockControl.ascx");
            BlockControls.ID = "BlockControl" + i.ToString();            
            Panel1.Controls.Add(BlockControls);
            ((Label)BlockControls.FindControl("Block1Num")).Text = Convert.ToString(i + 1);
                        ((DropDownList)BlockControls.FindControl("Block1DataNode")).DataSource = nodeItem_dt;
            ((DropDownList)BlockControls.FindControl("Block1DataNode")).DataTextField = "catName";
            ((DropDownList)BlockControls.FindControl("Block1DataNode")).DataValueField = "ctNodeId";
            ((DropDownList)BlockControls.FindControl("Block1DataNode")).DataBind();
            if (((DropDownList)BlockControls.FindControl("Block1DataNode")).Items.FindByValue(dsFormXml.Tables[1].Rows[i]["DataNode"].ToString()) != null )
            {
                ((DropDownList)BlockControls.FindControl("Block1DataNode")).SelectedValue = dsFormXml.Tables[1].Rows[i]["DataNode"].ToString();
            }
            else
            {
                ((DropDownList)BlockControls.FindControl("Block1DataNode")).SelectedValue = "0";
            }
            ((TextBox)BlockControls.FindControl("Block1Title")).Text = dsFormXml.Tables[1].Rows[i]["DataRemark"].ToString();
            ((TextBox)BlockControls.FindControl("Block1ContentLength")).Text = dsFormXml.Tables[1].Rows[i]["ContentLength"].ToString();
            ((CheckBox)BlockControls.FindControl("Block1IsPic")).Checked = BoolYN(dsFormXml.Tables[1].Rows[i]["IsPic"].ToString());
            ((CheckBox)BlockControls.FindControl("Block1IsTitle")).Checked = BoolYN(dsFormXml.Tables[1].Rows[i]["IsTitle"].ToString());
            ((CheckBox)BlockControls.FindControl("Block1IsExcerpt")).Checked = BoolYN(dsFormXml.Tables[1].Rows[i]["IsExcerpt"].ToString());
            ((CheckBox)BlockControls.FindControl("Block1IsPostDate")).Checked = BoolYN(dsFormXml.Tables[1].Rows[i]["IsPostDate"].ToString());
            ((RadioButton)BlockControls.FindControl("Block1Type1")).Checked = BoolRadio("Style1", dsFormXml.Tables[1].Rows[i]["ShowStyle"].ToString());
            ((RadioButton)BlockControls.FindControl("Block1Type2")).Checked = BoolRadio("Style2", dsFormXml.Tables[1].Rows[i]["ShowStyle"].ToString());
            ((RadioButton)BlockControls.FindControl("Block1Type3")).Checked = BoolRadio("Style3", dsFormXml.Tables[1].Rows[i]["ShowStyle"].ToString());
            if (((RadioButton)BlockControls.FindControl("Block1Type3")).Checked)
            {
                ((Panel)BlockControls.FindControl("Block1PanelSqlTop")).CssClass = "s2";
            }
            else
            {
                ((Panel)BlockControls.FindControl("Block1PanelSqlTop")).CssClass = "s1";
            }

            ((TextBox)BlockControls.FindControl("Block1SqlTop")).Text =dsFormXml.Tables[1].Rows[i]["SqlTop"].ToString();
            ((Button)BlockControls.FindControl("Block1Up")).Click += new EventHandler(ButtonUp_Click); 
            ((Button)BlockControls.FindControl("Block1Down")).Click +=new EventHandler(ButtonDown_Click);
            if (i == 0)
            {
                ((Button)BlockControls.FindControl("Block1Up")).CssClass = "s2";
            }
            else if(i==n-1)
            {
                ((Button)BlockControls.FindControl("Block1Down")).CssClass = "s2";
            }
            ((RadioButton)BlockControls.FindControl("Block1Type1")).InputAttributes.Add("onClick", "javascript: if (this.checked) { var obj = document.getElementById('BlockControl" + i.ToString() + "_Block1PanelSqlTop'); obj.className='s1';}");
            ((RadioButton)BlockControls.FindControl("Block1Type2")).InputAttributes.Add("onClick", "javascript: if (this.checked) { var obj = document.getElementById('BlockControl" + i.ToString() + "_Block1PanelSqlTop'); obj.className='s1';}");
            ((RadioButton)BlockControls.FindControl("Block1Type3")).InputAttributes.Add("onClick", "javascript: if (this.checked) { var obj = document.getElementById('BlockControl"+i.ToString()+"_Block1PanelSqlTop'); obj.className='s2';}");           
        }
        XmlFile.loadXmlFile(FilePath).Dispose();
       
    }

    protected void ButtonUp_Click(object sender, EventArgs e)
    {
        if (Session["Index"] == null)
        {
            return;
        }
        int x = Convert.ToInt32(Session["Index"]) - 2;
        int y = Convert.ToInt32(Session["Index"]);
        ShowData(x, y);
       
    }

    protected void ButtonDown_Click(object sender, EventArgs e)
    {
        if (Session["Index"] == null)
        {
            return;
        }
        int x = Convert.ToInt32(Session["Index"]) - 1;
        int y = Convert.ToInt32(Session["Index"]) + 1;
        ShowData(x, y);
    }
    //移動資料
    protected void ShowData(int x,int y)
    {
        DataSet ds = BlockControlData(x, y);
        UserControl BlockControls;       
        int i = x;
        for (int j = x; j < y; j++)
        {
            if (j == x)
            {
                i = 1;
            }
            else
            {
                i = 0;
            }
            BlockControls = (UserControl)Panel1.FindControl("BlockControl" + j.ToString());
            DataRow dr = ds.Tables[1].Rows[i];

            ((TextBox)BlockControls.FindControl("Block1Title")).Text = dr["DataRemark"].ToString();
            ((DropDownList)BlockControls.FindControl("Block1DataNode")).SelectedValue = dr["DataNode"].ToString();
            ((TextBox)BlockControls.FindControl("Block1SqlTop")).Text = dr["SqlTop"].ToString();
            ((CheckBox)BlockControls.FindControl("Block1IsTitle")).Checked = (bool)dr["IsTitle"];
            ((CheckBox)BlockControls.FindControl("Block1IsPic")).Checked = (bool)dr["IsPic"];
            ((CheckBox)BlockControls.FindControl("Block1IsPostDate")).Checked = (bool)dr["IsPostDate"];
            ((CheckBox)BlockControls.FindControl("Block1IsExcerpt")).Checked = (bool)dr["IsExcerpt"];
            ((TextBox)BlockControls.FindControl("Block1ContentLength")).Text = dr["ContentLength"].ToString();
            ((RadioButton)BlockControls.FindControl("Block1Type1")).Checked = (bool)dr["Type1"];
            ((RadioButton)BlockControls.FindControl("Block1Type2")).Checked = (bool)dr["Type2"];
            ((RadioButton)BlockControls.FindControl("Block1Type3")).Checked = (bool)dr["Type3"];
            if (((RadioButton)BlockControls.FindControl("Block1Type3")).Checked)
            {
                ((Panel)BlockControls.FindControl("Block1PanelSqlTop")).CssClass = "s2";
            }
            else
            {
                ((Panel)BlockControls.FindControl("Block1PanelSqlTop")).CssClass = "s1";
            }
        }
        ds.Dispose();       
        BlockControlData(x, y).Dispose();
        System.GC.Collect();
    }

    protected string savePath()
    {
        
        string message = "";
        message = Session["GipDsdPath"] + "\\xdmp" + Session["TreeRootId"].ToString() + ".xml"; if (!File.Exists(message))
        {
            message = "-1";
        }
        
        return message; 
    }

    protected bool BoolYN(string YN)
    {
        return (YN.ToUpper() == "Y") ? true : false;

    }
    protected bool BoolRadio(string radiovalue, string style)
    {        
       return (radiovalue.ToUpper() == style.ToUpper()) ? true : false;       
    }

    protected void saveXmlFile()
    {
        updateData().WriteXml(savePath());
    }
    protected void Next_Click(object sender, EventArgs e)
    {
        saveXmlFile();
        Response.Redirect("finish.aspx");
    }
    
    protected void Save_Click(object sender, EventArgs e)
    {
        saveXmlFile();
    }
   

    protected DataSet BlockControlData(int x, int y)
    {
         DataSet ds = new DataSet();
         CreateTable createViewTable = new CreateTable();
         ds.Tables.Add(createViewTable.createMyDataSet());
         ds.Tables.Add(createViewTable.createDataSet());        

         UserControl BlockControls;

        int ControlCount = y;
        
        for (int i = x; i < ControlCount; i++)
        {
            BlockControls = (UserControl)Panel1.FindControl("BlockControl" + i.ToString());
           DataRow dr = ds.Tables[1].NewRow();

            dr["DataRemark"] = ((TextBox)BlockControls.FindControl("Block1Title")).Text;
            dr["DataNode"] = ((DropDownList)BlockControls.FindControl("Block1DataNode")).SelectedValue;
            dr["SqlTop"] = ((TextBox)BlockControls.FindControl("Block1SqlTop")).Text;
            dr["IsTitle"] = ((CheckBox)BlockControls.FindControl("Block1IsTitle")).Checked;
            dr["IsPic"] = ((CheckBox)BlockControls.FindControl("Block1IsPic")).Checked;
            dr["IsPostDate"] = ((CheckBox)BlockControls.FindControl("Block1IsPostDate")).Checked;
            dr["IsExcerpt"] = ((CheckBox)BlockControls.FindControl("Block1IsExcerpt")).Checked;
            dr["ContentLength"] = ((TextBox)BlockControls.FindControl("Block1ContentLength")).Text;
            dr["Type1"] = ((RadioButton)BlockControls.FindControl("Block1Type1")).Checked;
            dr["Type2"] = ((RadioButton)BlockControls.FindControl("Block1Type2")).Checked;
            dr["Type3"] = ((RadioButton)BlockControls.FindControl("Block1Type3")).Checked;

            ds.Tables[1].Rows.Add(dr);
        }       
        createViewTable.createMyDataSet().Dispose();
        createViewTable.createDataSet().Dispose();        
        return ds;
    }
    //change
    public DataSet updateData()
    {
        DataSet OldData = XmlFile.loadXmlFile(savePath());       
        DataSet NewData = BlockControlData(0, Panel1.Controls.Count-1);
        int n = OldData.Tables[1].Rows.Count;
        
        for (int i = 0; i < n; i++)
        {
            DataRow OldDr = OldData.Tables[1].Rows[i];
            DataRow NewDr = NewData.Tables[1].Rows[i];

            OldDr["DataRemark"] = NewDr["DataRemark"];
            OldDr["DataNode"] = NewDr["DataNode"];
            OldDr["SqlTop"] = NewDr["SqlTop"];
            OldDr["IsTitle"] =boolYN((bool)NewDr["IsTitle"]);
            OldDr["IsPic"] = boolYN((bool)NewDr["IsPic"]);
            OldDr["IsPostDate"] = boolYN((bool)NewDr["IsPostDate"]);
            OldDr["IsExcerpt"] = boolYN((bool)NewDr["IsExcerpt"]);
            OldDr["ShowStyle"] =style((bool)NewDr["Type1"],(bool)NewDr["Type2"],(bool)NewDr["Type3"]);
             OldDr["ContentLength"] = NewDr["ContentLength"];
        }
        return OldData;
    }
    
    public string style(bool radiobutton1, bool radiobutton2, bool radiobutton3)
    {
        string message = "";
        if (radiobutton3)
        {
            message = "Style3";
        }
        else if (radiobutton2)
        {
            message = "Style2";
        }
        else
        {
            message = "Style1";
        }
        return message;
    }
    public string boolYN(bool CheckedValue)
    {
        return (CheckedValue) ? "Y" : "N";        
         
    }
    protected void Back_Click(object sender, EventArgs e)
    {
        Response.Redirect("./gip/web/step3.aspx");
    }
    protected void Home_Click(object sender, EventArgs e)
    {
        Response.Redirect("./index.aspx");
    }
}
