using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Drawing.Imaging;
using System.Drawing;
using System.Data.SQLite;
using System.Text;
using System.Data;

public partial class fullimage: System.Web.UI.Page
{
    protected string fileName="";
    
    protected void Page_Load(object sender, EventArgs e)
    {
	if(Request["fileName"] != null && Request["fileName"] !=""){
    		fileName=Request["fileName"];	
	}
    }
}