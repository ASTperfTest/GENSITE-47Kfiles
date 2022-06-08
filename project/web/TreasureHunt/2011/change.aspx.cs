using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class TreasureHunt_2011_change : System.Web.UI.Page
{
    protected string flashfile = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["memID"] == null || Session["changeTreasureId"] == null || string.IsNullOrEmpty(Session["changeTreasureId"].ToString()))
        {
            Response.StatusCode = 403;
            Response.End();
        }
        string meid = Session["memID"].ToString();
        string treasureId = Session["changeTreasureId"].ToString();
        flashfile = "/treasurehunt/2011/" + treasureId + ".swf";
    }
}