using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class reset : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string _key = Request.QueryString["key"];
        string _login_id = Request.Form["login_id"];
        string _game_key = Request.Form["game_key"];

        if (!FishBowlUtil.isMatchedKey(_login_id, _game_key, _key))
        {
            Response.Write("KeyError");
        }
        else
        {
            if (_login_id.Length >= 0 && _game_key.Length > 0)
            {
                FishBowl fb = new FishBowl();
                int account_id = fb.isValidKey(_login_id, _game_key);

                // 檢查key是否正確
                if (account_id > 0)
                {
                    int money = System.Convert.ToInt32(Request.Form["money"]);
                    double water_status = System.Convert.ToDouble(Request.Form["water_status"]);
                    fb.update_enviroment(account_id, money, water_status);
                    fb.reset_data(account_id);
                    
                    // 寫入完成，輸出資料
                    Response.Write("done");
                }
                else
                {
                    Response.Write("ErrorKey");
                }

                fb = null;
            }
            else
            {
                Response.Write("error");
            }
        }
    }
}
