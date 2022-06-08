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

public partial class Edit_Class : System.Web.UI.Page
{
    public Setting dbconfig = new Setting();
    public string class_id;
    protected void Page_Load(object sender, EventArgs e)
    {

        
        SqlDataSource1.ConnectionString = dbconfig.ConnectionSettings();
        SqlDataSource1.SelectCommand = "select * from type where datalevel = 1 order by sortvalue asc";

       
       


       
    }

   


    protected void SqlDataSource1_Selecting(object sender, SqlDataSourceSelectingEventArgs e)
    {

    }


}
