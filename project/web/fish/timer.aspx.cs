using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Threading;

public partial class timer : System.Web.UI.Page
{
    private string _login_id = "";
    private string _game_key = "";
    private string _key = "";

    protected void Page_Load(object sender, EventArgs e)
    {
		//Thread.Sleep(10000);

        _login_id = Request.Form["login_id"];
        _game_key = Request.Form["game_key"];
        _key = Request.QueryString["key"];

        if(FishBowlUtil.isMatchedKey(_login_id, _game_key, _key))
        {
            string sql = "SELECT e.create_time FROM fishbowl_environment e, fishbowl_account a WHERE e.account_id = a.id AND a.login_id=@LOGIN_ID";

            using(SqlConnection conn = SQLDB.createConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            Response.Write(FishBowlUtil.getTimeDelta(System.Convert.ToDateTime(dr["create_time"])));
                        }
                        else
                        {
                            Response.Write("0");
                        }
                    }
                }
            }
        }
        else
        {
            Response.Write ("0");
        }
    }
}
