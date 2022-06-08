using System;
using System.IO;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Xml;

public partial class question : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        // game_key
        string key = Request.QueryString["key"];

        // 傳送GameRank(遊戲等級)和QAcount(遊戲數目)這兩個參數過去
        string GameRank = Request.Form["GameRank"];
        string QAcount = Request.Form["QAcount"];

        string login_id = Request.Form["login_id"];
        string game_key = Request.Form["game_key"];

        if (game_key != null && GameRank != null && QAcount != null)
        {
            Farm2009 farm = new Farm2009(login_id, game_key);

            if (!farm.isMatchedKey(key))
            {
                Response.Write("KeyError");
            }
            else
            {
                // 隨機從所有符合(GameRank)的題目，取出(QAcount)個題目，並以XML回傳
                string question = farm.getQuestion(System.Convert.ToInt32(GameRank), System.Convert.ToInt32(QAcount));
                Response.ContentType = "text/xml";
                Response.Write(question);
            }
            farm = null;
        }
        else
        {
            Response.Write("KeyError");
        }

    }


}
