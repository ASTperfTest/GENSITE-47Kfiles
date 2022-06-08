using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;

public partial class TreasureHunt_membercaptchat : System.Web.UI.Page
{
    protected string treasreLogId = "";
    protected string targetpage = "";
    protected string pageParam = "";
    protected string guid = "";
    protected string refereurl = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null || Session["memID"].ToString().Trim() == "" || Session["treasureLogId"] == null || Session["treasureLogId"].ToString().Trim() == "")
        {
            Response.StatusCode = 403;
            Response.Write("403 存取被拒;");
            Response.End();
        }
		Session.Remove("treasureLogId");
        guid = Guid.NewGuid().ToString("N");
        treasreLogId = WebUtility.GetStringParameter("treasrelogid", string.Empty).ToLower();
        targetpage = WebUtility.GetStringParameter("targetpage", string.Empty).ToLower();
        pageParam = WebUtility.GetStringParameter("documentinfo", string.Empty).ToLower();
        refereurl = WebUtility.GetStringParameter("refereurl", string.Empty).ToLower();
    }
}
