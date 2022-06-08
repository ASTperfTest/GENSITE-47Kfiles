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

public partial class Finish : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string x = Session["User_id"].ToString();
        string url = System.Configuration.ConfigurationManager.AppSettings["Browserpath"].ToString();
        string id = x;

        browser.Attributes["onclick"] = "window.open('" + url + id + "')";
        Session.Remove("check");
        if (Session["URL"] == null)
        {

            modify.Visible = true;

        }
        else
        {
            modify.Visible = false;
            Session.Remove("URL");
        }
    }
    protected void modify_Click(object sender, EventArgs e)
    {
        Response.Redirect("../GipEdit/CtNodeTList.asp?CtRootID=4");
    }
    protected void gototop_Click(object sender, EventArgs e)
    {
        Response.Redirect("./index.aspx");
    }
}
