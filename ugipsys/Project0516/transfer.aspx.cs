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

public partial class transfer : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        
      string id = Request.QueryString["DBUSERID"].ToString();
	    Session.Add("Name",id);
	
	    string dPath = Request.QueryString["dGipDsdPath"].ToString(); 
      Session.Add("GipDsdPath",dPath );

      string wpath = Request.QueryString["wPublicPath"].ToString();
      Session.Add("PublicPath", wpath);
	    
      Response.Redirect("index.aspx?id="+id + "&dGipDsdPath="+dPath);
    }
}
