using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.Text;
using System.Security.Cryptography;

public partial class save : System.Web.UI.Page
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

                    double anemone_health_status = System.Convert.ToDouble(Request.Form["anemone_health_status"]);
                    double anemone_full_status = System.Convert.ToDouble(Request.Form["anemone_full_status"]);
                    int anemone_birth_secs = System.Convert.ToInt32(Request.Form["anemone_birth_secs"]);
                    int anemone_is_newborn = System.Convert.ToInt32(Request.Form["anemone_is_newborn"]);
                    if (anemone_birth_secs > 0)
                    {
                        fb.update_fish(account_id, 1, anemone_health_status, anemone_full_status, anemone_birth_secs, anemone_is_newborn);
                    }

                    double bream_health_status = System.Convert.ToDouble(Request.Form["bream_health_status"]);
                    double bream_full_status = System.Convert.ToDouble(Request.Form["bream_full_status"]);
                    int bream_birth_secs = System.Convert.ToInt32(Request.Form["bream_birth_secs"]);
                    int bream_is_newborn = System.Convert.ToInt32(Request.Form["bream_is_newborn"]);
                    if (bream_birth_secs > 0)
                    {
                        fb.update_fish(account_id, 2, bream_health_status, bream_full_status, bream_birth_secs, bream_is_newborn);
                    }

                    double clownfish_health_status = System.Convert.ToDouble(Request.Form["clownfish_health_status"]);
                    double clownfish_full_status = System.Convert.ToDouble(Request.Form["clownfish_full_status"]);
                    int clownfish_birth_secs = System.Convert.ToInt32(Request.Form["clownfish_birth_secs"]);
                    int clownfish_is_newborn = System.Convert.ToInt32(Request.Form["clownfish_is_newborn"]);
                    if (clownfish_birth_secs > 0)
                    {
                        fb.update_fish(account_id, 3, clownfish_health_status, clownfish_full_status, clownfish_birth_secs, clownfish_is_newborn);
                    }

                    double hippocampus_health_status = System.Convert.ToDouble(Request.Form["hippocampus_health_status"]);
                    double hippocampus_full_status = System.Convert.ToDouble(Request.Form["hippocampus_full_status"]);
                    int hippocampus_birth_secs = System.Convert.ToInt32(Request.Form["hippocampus_birth_secs"]);
                    int hippocampus_is_newborn = System.Convert.ToInt32(Request.Form["hippocampus_is_newborn"]);
                    if (hippocampus_birth_secs > 0)
                    {
                        fb.update_fish(account_id, 4, hippocampus_health_status, hippocampus_full_status, hippocampus_birth_secs, hippocampus_is_newborn);
                    }
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
