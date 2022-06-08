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

public partial class download : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
	if(Request["fileName"]!=null && Request["fileName"]!=""){
		DownloadFile(Request["fileName"]);
	}
    }
    
    private void DownloadFile(string fileName)
    {
       string filePath="/logoselection/images/matchlogo/"+fileName;
       string browserName=Request.Browser.Browser.ToLower();
       Response.Expires=0;
       Response.CacheControl="no-cache";
       Response.AddHeader("Pragma","no-cache");	
       Response.ContentType="application/save-as";
       if(browserName.Equals("ie")){
		Response.AddHeader("content-disposition","attachment;filename="+Server.UrlPathEncode(fileName));
	}else{
		Response.AddHeader("content-disposition","attachment;filename="+filePath);
	}
        Response.TransmitFile(filePath);
    }
}