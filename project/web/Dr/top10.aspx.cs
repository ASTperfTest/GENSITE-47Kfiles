using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class top10 : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // game_key
        string key = Request.QueryString["key"];

        string login_id = Request.Form["login_id"];
        string game_key = Request.Form["game_key"];

        Farm2009 farm = new Farm2009(login_id, game_key);
        if (farm.isMatchedKey(key))
        {
            Response.ContentType = "text/xml";
            string strScore = farm.getTopScore(10);
            farm = null;

            Response.Write(strScore);
        }
        else
        {
            Response.Write("KeyError");
        }
    }
}
