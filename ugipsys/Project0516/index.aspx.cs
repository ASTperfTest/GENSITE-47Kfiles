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

public partial class index : System.Web.UI.Page
{
  public Setting dbconfig = new Setting();
  public int recordcount;
  public Boolean isAdmin = false;

  string ssoID = string.Empty;

  protected void Page_Load(object sender, EventArgs e)
  {
      txtSubject.Attributes.Add("onkeypress", "if( event.keyCode == 13 ) {" + ClientScript.GetPostBackEventReference(btnSearch, "") + "}");

      if (Request.QueryString["id"] != null)
      {
          string userid = Request.QueryString["id"].ToString();
          Session.Add("Name", userid);
      }


    //2011/09/13
    //bob SSO ID, 主題館中未開放的目錄或文章，必須要有sso id才可以瀏覽
    //重要!!：
    //在主題館管理列表(/project0516/index.aspx) 及 
    //文件列表(GipEdit/DsdXMLList.asp) 都有這個需求，但由於這二支程式分別是asp及asp.net，無法相容，因此在修改這段code時，需同步修改另一個檔案
      using (SqlConnection connO = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString))
      {
          connO.Open();
          string strSQLO = @"            
            set nocount on
            delete sso where creation_datetime < getdate()-3;

            declare @guid   nvarchar(100)
            select @guid = [guid] from SSO where user_id = @userId
            if ( @guid is null)
            begin
	            set @guid = newid()
                insert into sso (user_id, [guid], CREATION_DATETIME)
                select @userId ,@guid ,getdate()
            end
            else
            begin
	            update sso set CREATION_DATETIME = getdate() 
	            where user_id = @userId and [guid] = @guid
            end
            select @guid as [guid]";


          SqlCommand cmd = connO.CreateCommand();
          cmd.CommandText = strSQLO;
          cmd.CommandType = CommandType.Text;
          cmd.Parameters.Add("@userId", SqlDbType.NVarChar).Value = "uAdm_" + Session["Name"];
          this.ssoID = cmd.ExecuteScalar().ToString();
      }
      //*******************************************************

      //舊主題館
      if (Request.QueryString["ctnodeid"] != null && Request.QueryString["type"] == "1")
      {
          using (SqlConnection connO = new SqlConnection(ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString))
          {
              connO.Open();
              string strSQLO = "UPDATE NodeInfo SET old_subject = @old_subject WHERE CtrootID = @CtrootID";
              SqlCommand cmd = new SqlCommand(strSQLO, connO);
              cmd.Parameters.Add("@old_subject", SqlDbType.Char).Value = "Y";
              cmd.Parameters.Add("@CtrootID", SqlDbType.Int).Value = Request.QueryString["ctnodeid"];
              cmd.ExecuteNonQuery();
              connO.Close();
          }
      }
      //舊主題館end
      using (SqlConnection conn = new SqlConnection(dbconfig.ConnectionSettings()))
      {
          SqlDataReader reader = null;
          try
          {
              string userRight = "SELECT ugrpID FROM InfoUser WHERE UserID = '" + Session["Name"].ToString() + "'";
              conn.Open();
              SqlCommand comm = new SqlCommand(userRight, conn);
              reader = comm.ExecuteReader();
              if (reader.Read())
              {
                  if (reader["ugrpID"].ToString().Contains("HTSD") || reader["ugrpID"].ToString().Contains("SysAdm"))
                  {
                      isAdmin = true;
                  }
              }
              reader.Close();
          }
          catch { }
      }

      getGridData();

      if (IsPostBack)
      {
          string ctnodeid = Request.QueryString["ctnodeid"];
          string type = Request.QueryString["type"];
          if (ctnodeid != "" && type == "1")
          {
              //Response.Write(ctnodeid);
          }
      }
      else
      {

          set_page(ddl_page, recordcount);
          show();
      }
      if (!isAdmin)
      {
          musicadd.Visible = false;
      }

  }

  private void getGridData()
  {
      SqlDataSource1.ConnectionString = dbconfig.ConnectionSettings();

      string strSQL = "", subjectName = txtSubject.Text.Trim();
      
      strSQL = @"select CatTreeRoot.*
                    ,nodeinfo.*
                    ,InfoUser.UserName
                    ,isnull(a.ViewCount ,0) ViewCount
                from  InfoUser 
                RIGHT OUTER JOIN NodeInfo ON InfoUser.UserID = NodeInfo.owner 
                LEFT OUTER JOIN CatTreeRoot ON NodeInfo.CtrootID = CatTreeRoot.CtRootID ";
	  strSQL += " LEFT JOIN ( ";
      strSQL += "Select ctRootId, sum(ViewCount) as ViewCount ";
	  strSQL += "	 from CounterForSubjectByDate ";
	  strSQL += "	 GROUP BY  ctRootId ";
      strSQL += " )a ON a.ctRootId = CatTreeRoot.ctrootid ";
	  strSQL += "where 1 = 1 ";
      strSQL += " and (old_subject = 'N' or old_subject is null ) ";
      if (subjectName != "") strSQL += " and CatTreeRoot.CtRootName like '%" + subjectName + "%' ";

      if (IsPublic.SelectedValue != "")
      {
          strSQL += " and CatTreeRoot.inUse = '" + IsPublic.SelectedValue + "' ";
      }      
      if (!isAdmin)
      {
          strSQL += "and nodeinfo.owner ='" + Session["Name"].ToString() + "' ";
      }
      else
      {
          strSQL += "and vGroup = 'XX' ";
      }

      strSQL += @"order by 
        CatTreeRoot.inUse desc
        ,(Case When NodeInfo.order_num is null Then 1 Else 0 End)
        , NodeInfo.order_num DESC
        , cattreeroot.CtRootID ";

      SqlDataSource1.SelectCommand = strSQL;
      GridView1.DataBind();


      SqlDataSource ds = GridView1.DataSourceObject as SqlDataSource;
      DataView view = ds.Select(new DataSourceSelectArguments()) as DataView;

      recordcount = view.Table.Rows.Count;
      datacount.Text = view.Table.Rows.Count.ToString();
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
    
    set_page(ddl_page, recordcount);
    GridView1.PageIndex = 0;
    next.Enabled = true;
    back.Enabled = true;
    GridView1.DataBind();
    show();
  }
  
  protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
  {
	//要增加欄位請注意index是否會影響到隱藏欄位 記得修改下方是否為admin的部分
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
      url += "&ssoID=" + ssoID;
      x.EnableViewState = false;
      x.Attributes["onclick"] = "window.open('" + url + "')";
      Button btn = (Button)e.Row.FindControl("Old");      
      btn.PostBackUrl = "index.aspx?ctnodeid=" + id + "&type=1";
      btn.OnClientClick = "if(confirm('注意!確定要封存主題館嗎?(搬遷至主題館封存區)')){location.href='" + btn.PostBackUrl + "';}else{window.event.returnValue=false;}";
    }
	  //判斷是否為admin
	  if (!isAdmin)
    {
		  e.Row.Cells[6].Visible = false; 
		  e.Row.Cells[7].Visible = false; 
		  e.Row.Cells[8].Visible = false; 
	  }
  }
  protected void btnSearch_Click(object sender, EventArgs e)
  {
      getGridData();
      set_page(ddl_page, recordcount);
      show();      
  }
}



