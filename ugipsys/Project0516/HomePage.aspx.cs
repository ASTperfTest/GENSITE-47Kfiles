using System;
using System.Data;
using System.Collections;
using System.Xml;
using System.Web.UI.WebControls;
using System.Text;
using System.IO;

public partial class HomePage : System.Web.UI.Page
{   
    protected void Page_Load(object sender, EventArgs e)
    {

             

        if (!IsPostBack)
        { 
            Session["TreeRootId"] = Session["User_id"].ToString();
            Session["ID"] = Session["Name"].ToString();          
	    FileExist();
        }

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
Response.Write(Page.Server.MapPath(PicPath));
        }

        Display(); 
    }
    protected void Display()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("javascript:");
        sb.Append("var objc = document.getElementById('Block1Type3');");
        sb.Append("if(objc.checked == true) {");
        sb.Append("var obj = document.getElementById('Block1PanelSqlTop'); ");
        sb.Append(" obj.className ='s2'; ");
        sb.Append("}");
        sb.Append("else {");
        sb.Append("var obj = document.getElementById('Block1PanelSqlTop'); ");
        sb.Append(" obj.className ='s1'; ");
        sb.Append("}");

        Block1Type1.InputAttributes.Add("onClick", sb.ToString());
        Block1Type2.InputAttributes.Add("onClick", sb.ToString());
        Block1Type3.InputAttributes.Add("onClick", sb.ToString());       
        sb.Replace("Block1", "Block2");

        Block2Type1.InputAttributes.Add("onClick", sb.ToString());
        Block2Type2.InputAttributes.Add("onClick", sb.ToString());
        Block2Type3.InputAttributes.Add("onClick", sb.ToString());
        sb.Replace("Block2", "Block3");

        Block3Type1.InputAttributes.Add("onClick", sb.ToString());
        Block3Type2.InputAttributes.Add("onClick", sb.ToString());
        Block3Type3.InputAttributes.Add("onClick", sb.ToString());
        sb.Replace("Block3", "Block4");

        Block4Type1.InputAttributes.Add("onClick", sb.ToString());
        Block4Type2.InputAttributes.Add("onClick", sb.ToString());
        Block4Type3.InputAttributes.Add("onClick", sb.ToString());
        sb.Replace("Block4", "Block5");

        Block5Type1.InputAttributes.Add("onClick", sb.ToString());
        Block5Type2.InputAttributes.Add("onClick", sb.ToString());
        Block5Type3.InputAttributes.Add("onClick", sb.ToString());       

    }


    protected void Save_Click(object sender, EventArgs e)
    {
        Display();   
        CreateXdmpXml createXmlFile = new CreateXdmpXml();
        createXmlFile.createXmlFile(BlockAllData(), savePath());
        xdmpListInfo();       
    }
    protected void Next_Click(object sender, EventArgs e)
    {
        Display();   
        xdmpListInfo();
        CreateXdmpXml createXmlFile = new CreateXdmpXml();
        createXmlFile.createXmlFile(BlockAllData(), savePath());        
//        Page.Server.Transfer("HomePageEdit.aspx");
	Response.Redirect("finish.aspx");
    }

    public DataTable createMyDataSet()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("MenuTree", typeof(string));
        dt.Columns.Add("MpStyle", typeof(string));
        return dt;
    }
    public DataTable createDataSet()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("DataLable", typeof(string));
        dt.Columns.Add("DataRemark", typeof(string));
        dt.Columns.Add("DataNode", typeof(string));
        dt.Columns.Add("SqlCondition", typeof(string));
        dt.Columns.Add("SqlOrderBy", typeof(string));
        dt.Columns.Add("SqlTop", typeof(string));
        dt.Columns.Add("WithContent", typeof(string));
        dt.Columns.Add("ContentData", typeof(string));
        dt.Columns.Add("ContentLength", typeof(string));
        dt.Columns.Add("IsTitle", typeof(bool));
        dt.Columns.Add("IsPic", typeof(bool));
        dt.Columns.Add("IsPostDate", typeof(bool));
        dt.Columns.Add("IsExcerpt", typeof(bool));
        dt.Columns.Add("Type1", typeof(bool));
        dt.Columns.Add("Type2", typeof(bool));
        dt.Columns.Add("Type3", typeof(bool));
        return dt;
    }

    public void TreeData(DataTable dt, string MenuTree, string MpStyle)
    {
        DataRow dr = dt.NewRow();
        dr["MenuTree"] = MenuTree.ToString();
        dr["MpStyle"] = MpStyle.ToString();
        dt.Rows.Add(dr);
    }

    public void BlockData(DataTable dt, string DataLabel, string DataRemark, string DataNode, string SqlCondition, string SqlOrderBy,
        string SqlTop, string WithContent,string ContentData, string ContentLength, bool IsTitle, bool IsPic, bool IsPostDate,
        bool IsExcerpt, bool Type1, bool Type2, bool Type3)
    {       
        DataRow dr = dt.NewRow();
        dr["DataLable"] = DataLabel.ToString();
        dr["DataRemark"] = DataRemark.ToString();
        dr["DataNode"] = DataNode.ToString();
        dr["SqlCondition"] = SqlCondition.ToString();
        dr["SqlOrderBy"] = SqlOrderBy.ToString();
        dr["SqlTop"] = SqlTop.ToString();
        dr["WithContent"] = WithContent.ToString();
        dr["ContentData"] = ContentData.ToString();
        dr["ContentLength"]= ContentLength.ToString();        
        dr["IsTitle"] = (bool)IsTitle;
        dr["IsPic"] = (bool)IsPic;
        dr["IsPostDate"] = (bool)IsPostDate;
        dr["IsExcerpt"] = (bool)IsExcerpt;
        dr["Type1"] = (bool)Type1;
        dr["Type2"] = (bool)Type2;
        dr["Type3"] = (bool)Type3;
        dt.Rows.Add(dr);
        dt.Dispose();       
    }

    public DataSet BlockAllData()
    {
        DataSet ds = new DataSet();
        ds.Tables.Add(createMyDataSet());
        ds.Tables.Add( createDataSet());

        TreeData(ds.Tables[0], Session["TreeRootId"].ToString(), "Style1");
        
        BlockData(ds.Tables[1], "BlockA", Block1Title.Text, Block1DataNode.SelectedValue, "", "", Block1SqlTop.Text, "N", "N",
             Block1ContentLength.Text, Block1IsTitle.Checked, Block1IsPic.Checked, Block1IsPostDate.Checked, Block1IsExcerpt.Checked, Block1Type1.Checked, Block1Type2.Checked, Block1Type3.Checked);

        BlockData(ds.Tables[1], "BlockB", Block2Title.Text, Block2DataNode.SelectedValue, "", "", Block2SqlTop.Text, "N", "N",
              Block2ContentLength.Text, Block2IsTitle.Checked, Block2IsPic.Checked, Block2IsPostDate.Checked, Block2IsExcerpt.Checked, Block2Type1.Checked, Block2Type2.Checked, Block2Type3.Checked);

        BlockData(ds.Tables[1], "BlockC", Block3Title.Text, Block3DataNode.SelectedValue, "", "", Block3SqlTop.Text, "N", "N",
             Block3ContentLength.Text, Block3IsTitle.Checked, Block3IsPic.Checked, Block3IsPostDate.Checked, Block3IsExcerpt.Checked, Block3Type1.Checked, Block3Type2.Checked, Block3Type3.Checked);

        BlockData(ds.Tables[1], "BlockD", Block4Title.Text, Block4DataNode.SelectedValue, "", "", Block4SqlTop.Text, "N", "N",
             Block4ContentLength.Text, Block4IsTitle.Checked, Block4IsPic.Checked, Block4IsPostDate.Checked, Block4IsExcerpt.Checked, Block4Type1.Checked, Block4Type2.Checked, Block4Type3.Checked);

        BlockData(ds.Tables[1], "BlockE", Block5Title.Text, Block5DataNode.SelectedValue, "", "", Block5SqlTop.Text, "N", "N",
             Block5ContentLength.Text, Block5IsTitle.Checked, Block5IsPic.Checked, Block5IsPostDate.Checked, Block5IsExcerpt.Checked, Block5Type1.Checked, Block5Type2.Checked, Block5Type3.Checked);

        ds.Dispose();         
        return ds;        
    }
    
   //新增資料至資料庫
    public void xdmpListInfo()
    {
        HomePageSql createXdmpList = new HomePageSql();
        createXdmpList.createXdmpList(Convert.ToInt32(Session["TreeRootId"].ToString()), System.DBNull.Value, Session["ID"].ToString(), System.DBNull.Value);       
    }
    //檔案存放路徑
    protected string savePath()
    {
        if (Session["TreeRootId"] == null)
        {
            Session["TreeRootId"] = "";
        }
 	String dPath = Session["GipDsdPath"] + "\\xdmp" + Session["TreeRootId"].ToString() + ".xml" ;
	return dPath;
        //return Page.Server.MapPath("xdmp" + Session["TreeRootId"].ToString() + ".xml");
        //return Page.Server.MapPath("/site/" + session("mySiteID") + "/GipDSD/xdmp" + Session["TreeRootId"].ToString() + ".xml");
    }
   
    protected void Block1Data(DataSet ds, int A)
    {       
        Block1DataNode.SelectedValue = ds.Tables[1].Rows[A]["DataNode"].ToString();
        Block1Title.Text = ds.Tables[1].Rows[A]["DataRemark"].ToString();
        Block1ContentLength.Text = ds.Tables[1].Rows[A]["ContentLength"].ToString();
        Block1IsPic.Checked = (bool)ds.Tables[1].Rows[A]["IsPic"];
        Block1IsExcerpt.Checked = (bool)ds.Tables[1].Rows[A]["IsExcerpt"];
        Block1IsPostDate.Checked = (bool)ds.Tables[1].Rows[A]["IsPostDate"];
        Block1Type1.Checked = (bool)ds.Tables[1].Rows[A]["Type1"];
        Block1Type2.Checked = (bool)ds.Tables[1].Rows[A]["Type2"];
        if ((bool)ds.Tables[1].Rows[A]["Type3"])
        {
            Block1PanelSqlTop.CssClass = "s2";
        }
        Block1Type3.Checked = (bool)ds.Tables[1].Rows[A]["Type3"];
        Block1SqlTop.Text = ds.Tables[1].Rows[A]["SqlTop"].ToString();
       
    }
    protected void Block2Data(DataSet ds, int B)
    {
         
        Block2DataNode.SelectedValue = ds.Tables[1].Rows[B]["DataNode"].ToString();
        Block2Title.Text = ds.Tables[1].Rows[B]["DataRemark"].ToString();
        Block2ContentLength.Text = ds.Tables[1].Rows[B]["ContentLength"].ToString();
        Block2IsPic.Checked = (bool)ds.Tables[1].Rows[B]["IsPic"];
        Block2IsTitle.Checked = (bool)ds.Tables[1].Rows[B]["IsTitle"];
        Block2IsExcerpt.Checked = (bool)ds.Tables[1].Rows[B]["IsExcerpt"];
        Block2IsPostDate.Checked = (bool)ds.Tables[1].Rows[B]["IsPostDate"];
        Block2Type1.Checked = (bool)ds.Tables[1].Rows[B]["Type1"];
        Block2Type2.Checked = (bool)ds.Tables[1].Rows[B]["Type2"];
        if ((bool)ds.Tables[1].Rows[B]["Type3"])
        {
            Block2PanelSqlTop.CssClass = "s2";
        }
        Block2Type3.Checked = (bool)ds.Tables[1].Rows[B]["Type3"];
        Block2SqlTop.Text = ds.Tables[1].Rows[B]["SqlTop"].ToString();

        
      
    }
    protected void Block3Data(DataSet ds, int C)
    {
        
        Block3DataNode.SelectedValue = ds.Tables[1].Rows[C]["DataNode"].ToString();
        Block3Title.Text = ds.Tables[1].Rows[C]["DataRemark"].ToString();
        Block3ContentLength.Text = ds.Tables[1].Rows[C]["ContentLength"].ToString();
        Block3IsPic.Checked = (bool)ds.Tables[1].Rows[C]["IsPic"];
        Block3IsTitle.Checked = (bool)ds.Tables[1].Rows[C]["IsTitle"];
        Block3IsExcerpt.Checked = (bool)ds.Tables[1].Rows[C]["IsExcerpt"];
        Block3IsPostDate.Checked = (bool)ds.Tables[1].Rows[C]["IsPostDate"];
        Block3Type1.Checked = (bool)ds.Tables[1].Rows[C]["Type1"];
        Block3Type2.Checked = (bool)ds.Tables[1].Rows[C]["Type2"];
        if ((bool)ds.Tables[1].Rows[C]["Type3"])
        {
            Block3PanelSqlTop.CssClass = "s2";
        }
        Block3Type3.Checked = (bool)ds.Tables[1].Rows[C]["Type3"];
        Block3SqlTop.Text = ds.Tables[1].Rows[C]["SqlTop"].ToString();

        
    }
    protected void Block4Data(DataSet ds, int D)
    {
         
        Block4DataNode.SelectedValue = ds.Tables[1].Rows[D]["DataNode"].ToString();
        Block4Title.Text = ds.Tables[1].Rows[D]["DataRemark"].ToString();
        Block4ContentLength.Text = ds.Tables[1].Rows[D]["ContentLength"].ToString();
        Block4IsPic.Checked = (bool)ds.Tables[1].Rows[D]["IsPic"];
        Block4IsTitle.Checked = (bool)ds.Tables[1].Rows[D]["IsTitle"];
        Block4IsExcerpt.Checked = (bool)ds.Tables[1].Rows[D]["IsExcerpt"];
        Block4IsPostDate.Checked = (bool)ds.Tables[1].Rows[D]["IsPostDate"];
        Block4Type1.Checked = (bool)ds.Tables[1].Rows[D]["Type1"];
        Block4Type2.Checked = (bool)ds.Tables[1].Rows[D]["Type2"];
        if ((bool)ds.Tables[1].Rows[D]["Type3"])
        {
            Block4PanelSqlTop.CssClass = "s2";
        }
        Block4Type3.Checked = (bool)ds.Tables[1].Rows[D]["Type3"];
        Block4SqlTop.Text = ds.Tables[1].Rows[D]["SqlTop"].ToString();

       

    }
    protected void Block5Data(DataSet ds, int E)
    {
        
        Block5DataNode.SelectedValue = ds.Tables[1].Rows[E]["DataNode"].ToString();
        Block5Title.Text = ds.Tables[1].Rows[E]["DataRemark"].ToString();
        Block5ContentLength.Text = ds.Tables[1].Rows[E]["ContentLength"].ToString();
        Block5IsPic.Checked = (bool)ds.Tables[1].Rows[E]["IsPic"];
        Block5IsTitle.Checked = (bool)ds.Tables[1].Rows[E]["IsTitle"];
        Block5IsExcerpt.Checked = (bool)ds.Tables[1].Rows[E]["IsExcerpt"];
        Block5IsPostDate.Checked = (bool)ds.Tables[1].Rows[E]["IsPostDate"];
        Block5Type1.Checked = (bool)ds.Tables[1].Rows[E]["Type1"];
        Block5Type2.Checked = (bool)ds.Tables[1].Rows[E]["Type2"];
        if ((bool)ds.Tables[1].Rows[E]["Type3"])
        {
            Block5PanelSqlTop.CssClass = "s2";
        }
        Block5Type3.Checked = (bool)ds.Tables[1].Rows[E]["Type3"];
        Block5SqlTop.Text = ds.Tables[1].Rows[E]["SqlTop"].ToString();

     
    }

    protected void Block1Down_Click(object sender, EventArgs e)
    {
        DataSet ds = BlockAllData();
        Block1Data(ds, 1);
        Block2Data(ds, 0);
        ds.Dispose();
        Display();   
        System.GC.Collect();

    }   
    protected void Block2Down_Click(object sender, EventArgs e)
    {
        DataSet ds = BlockAllData();
        Block2Data(ds, 2);
        Block3Data(ds, 1);
        ds.Dispose();
        Display();   
        System.GC.Collect();
    }   
    protected void Block3Down_Click(object sender, EventArgs e)
    {
        DataSet ds = BlockAllData();
        Block3Data(ds, 3);
        Block4Data(ds, 2);
        ds.Dispose();
        Display();   
        System.GC.Collect();
    }   
    protected void Block4Down_Click(object sender, EventArgs e)
    {
        DataSet ds = BlockAllData();
        Block4Data(ds, 4);
        Block5Data(ds, 3);
        ds.Dispose();
        Display();   
        System.GC.Collect();
    }   

       protected void FileExist()
    {
        string message = savePath();
        if (File.Exists(message))
            Response.Redirect("HomePageEdit.aspx");
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
