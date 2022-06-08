using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class user_click : System.Web.UI.Page
{
    private string _login_id = "";
    private string _game_key = "";
    private int _fish_id = 0;
    private string _key = "";

    protected void Page_Load(object sender, EventArgs e)
    {
        _login_id = Request.Form["login_id"];
        _game_key = Request.Form["game_key"];
        _fish_id = System.Convert.ToInt32(Request.Form["fish_id"]);
        _key = Request.QueryString["key"];

        if (FishBowlUtil.isMatchedKey(_login_id, _game_key, _key))
        {
            if (_fish_id > 0)
            {
                // 將use_know 欄位標記為1，代表使用者已知道死亡或成功
                FishBowl fb = new FishBowl();
                fb.update_user_know(_fish_id);
                fb = null;
                Response.Write("OK");
            }
            else
            {
                Response.Write("NoUpdate");
            }
        }
        else
        {
            Response.Write("KeyError");
        }
    }
}
