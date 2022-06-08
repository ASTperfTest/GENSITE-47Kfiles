using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
 
public partial class scorerank: Page
{

    protected void Page_Load(object sender, EventArgs e)
    {
		Response.Redirect("../dr/top_20102.aspx");
    }
}
 