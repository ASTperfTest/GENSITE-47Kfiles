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
using System.Data.SqlClient;

public partial class old_index : System.Web.UI.Page
{
  public Setting dbconfig = new Setting();
  public int recordcount;
  public Boolean isAdmin = false;

  protected void Page_Load(object sender, EventArgs e)
  {
    if(Request.QueryString["id"] != null)
    {
	    string userid = Request.QueryString["id"].ToString();
	    Session.Add("Name", userid);
	  }
	  //舊主題館
	if(Request.QueryString["ctnodeid"] != null && Request.QueryString["type"] =="2")
	{
		SqlConnection connO = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString);
		connO.Open();
		string strSQLO = "UPDATE NodeInfo SET old_subject = @old_subject WHERE CtrootID = @CtrootID";
		SqlCommand cmd = new SqlCommand(strSQLO, connO);
        cmd.Parameters.Add("@old_subject", SqlDbType.Char).Value = "N";
		cmd.Parameters.Add("@CtrootID", SqlDbType.Int).Value = Request.QueryString["ctnodeid"];
		cmd.ExecuteNonQuery();
        connO.Close();
	}
	//舊主題館end
    SqlConnection conn = null;
    SqlDataReader reader = null;
    try {
      string userRight = "SELECT ugrpID FROM InfoUser WHERE UserID = '" + Session["Name"].ToString() + "'";
      conn = new SqlConnection(dbconfig.ConnectionSettings());
      conn.Open();
      SqlCommand comm = new SqlCommand(userRight, conn);
      reader = comm.ExecuteReader();
      if ( reader.Read() ) {
        if (reader["ugrpID"].ToString().Contains("HTSD") || reader["ugrpID"].ToString().Contains("SysAdm") ) {
          isAdmin = true;
        }
      }
      reader.Close();
    }
    catch ( Exception ex) {}
    finally { 
	  if(conn.State == ConnectionState.Open )
        conn.Close();
    }
            
    SqlDataSource1.ConnectionString = dbconfig.ConnectionSettings();
    
    string strSQL = "";
    //strSQL = "select * from cattreeroot RIGHT OUTER JOIN nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where 1 = 1 ";
	  //strSQL = "select * from cattreeroot left join nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where 1 = 1 ";
	  strSQL = "select CatTreeRoot.*,nodeinfo.*,InfoUser.UserName from  InfoUser RIGHT OUTER JOIN NodeInfo ON InfoUser.UserID = NodeInfo.owner LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID where 1 = 1 ";
	 strSQL += " and old_subject = 'Y' ";
    if (!isAdmin) 
    {
      strSQL += "and nodeinfo.owner ='" + Session["Name"].ToString() + "' ";
    }
    else 
    {
      strSQL += "and vGroup = 'XX' ";
    }
	  strSQL += "order by ";
	  strSQL += " (Case When NodeInfo.order_num is null Then 1 Else 0 End), NodeInfo.order_num DESC, cattreeroot.CtRootID ";
	
    SqlDataSource1.SelectCommand = strSQL;
    GridView1.DataBind();


    if (!IsPostBack)
    {      
      get_count();
      set_page(ddl_page, recordcount);
      show();      
    }
    else
    {
      string ctnodeid = Request.QueryString["ctnodeid"];
      string type = Request.QueryString["type"];      
      if (ctnodeid != "" && type == "1")
      {
        //Response.Write(ctnodeid);
      }
    }
  }

  protected void get_count()
  {
    string strSQL = "";
    SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings());
    SqlCommand cmd = null;

    try {
      //strSQL = "select count(*) from cattreeroot left join nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where 1 = 1 ";
	  strSQL = "select count(*) from cattreeroot RIGHT OUTER JOIN nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid where 1 = 1 ";
	  strSQL += " and old_subject = 'Y' ";
      if (!isAdmin) {
        strSQL += "and nodeinfo.owner ='" + Session["Name"].ToString() + "'";
      }
      else {
        strSQL += "and vGroup = 'XX' ";
      }    
      conn.Open();
      cmd = new SqlCommand(strSQL,conn);
      recordcount = Convert.ToInt32(cmd.ExecuteScalar());
      datacount.Text = Convert.ToString(recordcount);      
    }
    catch (Exception ex) { }
    finally {
      conn.Close();
    }        
  }

  protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
  {        
    GridView1.PageIndex = e.NewPageIndex;
    ddl_page.SelectedIndex = e.NewPageIndex;
    //GridView1.DataBind();
    show();
  }

  protected void back_Click(object sender, EventArgs e)
  {
    if (GridView1.PageIndex > 0) {
      GridView1.PageIndex -= 1;
      ddl_page.SelectedIndex = GridView1.PageIndex;
      //GridView1.DataBind();
      show();
    }
  }

  protected void next_Click(object sender, EventArgs e)
  {
    if ( GridView1.PageIndex < (GridView1.PageCount - 1) ) {
      GridView1.PageIndex += 1;
      ddl_page.SelectedIndex = GridView1.PageIndex;
      show();
    }
  }

  protected void show()
  {
    if (GridView1.PageIndex == 0) {
      back.Enabled = false;
      next.Enabled = true;
    }
    else if ( GridView1.PageIndex == (GridView1.PageCount - 1) ) {
      next.Enabled = false;
      back.Enabled = true;
    }
    else {
      back.Enabled = true;
      next.Enabled = true;
    }

    if (GridView1.PageCount == 0 || GridView1.PageCount == 1) {           
      next.Enabled = false;
    }
  }

  protected void set_page(DropDownList DDL, int total)
  {
    ddl_page.Items.Clear();
    int x;
    if(total % GridView1.PageSize == 0) {
      x = total / GridView1.PageSize;
    }
    else {
      x = (total / GridView1.PageSize) + 1;
    }
    for (int i = 0; i < x; i++) {            
      string page_value = Convert.ToString(i+1);
      ListItem li = new ListItem(page_value,page_value);
      DDL.Items.Add(li);
    }    
  }
  
  protected void ddl_page_SelectedIndexChanged(object sender, EventArgs e)  
  {
    GridView1.PageIndex = Convert.ToInt32(ddl_page.SelectedItem.Text) -1;
    ddl_page.SelectedIndex = Convert.ToInt32(ddl_page.SelectedItem.Text) - 1;
    show();
  }

  protected void pagesize_SelectedIndexChanged(object sender, EventArgs e)
  {
    GridView1.PageSize = Convert.ToInt32(pagesize.SelectedValue);
    ddl_page.Items.Clear();
    get_count();
    set_page(ddl_page, recordcount);
    GridView1.PageIndex = 0;
    next.Enabled = true;
    back.Enabled = true;
    GridView1.DataBind();
    show();
  }
  
  protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
  {
    if (e.Row.RowIndex != -1) {
      LinkButton x = (LinkButton)e.Row.FindControl("Namelink");
      DataRowView rowView = (DataRowView)e.Row.DataItem;

      string url = System.Configuration.ConfigurationManager.AppSettings["Browserpath"].ToString();
      string id = rowView["ctrootid"].ToString();
      string urlLink = rowView["url_link"].ToString();

      if (urlLink == "")
      {
        url = url + id;
      }
      else
      {
        url = urlLink;
      }

      x.Attributes["href"] = "/GipEdit/old_CtNodeTList.asp?id="+id;
	  //x.Attributes["onclick"] = "window.open('" + url + "')";
      Button btn = (Button)e.Row.FindControl("restore");      
      btn.PostBackUrl = "old_index.aspx?ctnodeid=" + id + "&type=2";
      btn.OnClientClick = "location.href='" + btn.PostBackUrl + "';";
    }
	
  }
}



