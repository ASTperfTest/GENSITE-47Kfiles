using System;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using log4net;

public class FishBowl
{
    private SqlConnection _conn;
    private SqlCommand cmd;
    private SqlDataReader dr;
    private string sql;
    private ILog log4 = LogManager.GetLogger("FishBowl");
    /**
     * Constructor
     */
	public FishBowl(){
        log4net.Config.XmlConfigurator.Configure();
    }

    /**
     * DB Connection
     */
    private SqlConnection conn
    {
        get
        {
            // 如果還有connection...
            if (_conn != null)
            {
                // 如果沒開...
                if (_conn.State == ConnectionState.Closed)
                {
                    _conn.Open();
                }
                return _conn;
            }
            else
            {
                // 回傳一個新的
                _conn = SQLDB.createConnection();
                return _conn;
            }
        }
    }

    /**
     * 更新環境資料
     */
    public void update_enviroment(int _account_id, int _money, double _water_status)
    {
        sql = "UPDATE fishbowl_environment SET water_status=@WATER_STATUS, money=@MONEY, last_modify=@LAST_MODIFY WHERE account_id=@ACCOUNT_ID";

        try
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@WATER_STATUS", SqlDbType.Float).Value = _water_status;
                cmd.Parameters.Add("@MONEY", SqlDbType.Int, 4).Value = _money;
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = _account_id;
                cmd.Parameters.Add("@LAST_MODIFY", SqlDbType.DateTime).Value = DateTime.Now;
                cmd.ExecuteNonQuery();
            }
        }
        catch(Exception e)
        {
            log4.Error(e.ToString());
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    /**
     * 更新魚的資料
     */
    public void update_fish(int _account_id, int _fish_type, double _health_status, double _full_status, int _birth_sec, int _is_newborn)
    {
        sql = "";

        int last_fish_id = 0;

        // 傳入的生長天數大於零...先檢查該魚是否還沒長出來(create_time = null)

        try
        {
            if (_birth_sec > 0)
            {
                if (_is_newborn == 1)
                {
                    // 新生的魚
                    sql = "INSERT INTO fishbowl_pets(account_id, fish_type, create_time, status) VALUES(@ACCOUNT_ID, @FISH_TYPE, @CREATE_TIME, 1)";

                    using (cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = _account_id;
                        cmd.Parameters.Add("@FISH_TYPE", SqlDbType.TinyInt, 4).Value = _fish_type;
                        cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "Select @@Identity";
                        object o = cmd.ExecuteScalar();
                        if (!(o is DBNull))
                        {
                            last_fish_id = Convert.ToInt32(o);
                        }

                        string sql_add_environment_score = "INSERT INTO fishbowl_pets_score(fish_id, create_time) VALUES(@FISH_ID, @CREATE_TIME)";
                        cmd.CommandText = sql_add_environment_score;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = last_fish_id;
                        cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.ExecuteNonQuery();

                        // 紀錄
                        log(_account_id, "AddNewFish " + last_fish_id);
                    }
                }
                else
                {
                    sql = "SELECT id FROM fishbowl_pets WHERE account_id=@ACCOUNT_ID ANd fish_type=@FISH_TYPE AND create_time is NULL";

                    using (cmd = new SqlCommand(sql, conn))
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = _account_id;
                        cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = _fish_type;

                        using (dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                int _fish_id = System.Convert.ToInt32(dr["id"]);
                                dr.Close();

                                // 新魚，新增create_time資料
                                sql = "UPDATE fishbowl_pets SET create_time=@CREATE_TIME, status = 1 WHERE id=@FISH_ID";
                                cmd.CommandText = sql;
                                cmd.Parameters.Clear();
                                cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = _fish_id;
                                cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                dr.Close();
                                // 舊魚，更新資料
                                sql = "UPDATE fishbowl_pets SET health_status=@HEALTH_STATUS, full_status=@FULL_STATUS, last_modify=@LAST_MODIFY WHERE account_id=@ACCOUNT_ID AND fish_type=@FISH_TYPE";
                                cmd = new SqlCommand(sql, conn);
                                cmd.Parameters.Clear();
                                cmd.Parameters.Add("@HEALTH_STATUS", SqlDbType.Float).Value = _health_status;
                                cmd.Parameters.Add("@FULL_STATUS", SqlDbType.Float).Value = _full_status;
                                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = _account_id;
                                cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int).Value = _fish_type;
                                cmd.Parameters.Add("@LAST_MODIFY", SqlDbType.DateTime).Value = DateTime.Now;
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }
        catch (Exception e)
        {
            log4.Error(e.ToString());
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    /**
     * 檢查是否為有效之Game key
     */
    public int isValidKey(string _login_id, string _game_key)
    {
        int is_valid = 0;
        
        string sql_check_key = "SELECT * FROM fishbowl_account WHERE login_id=@LOGIN_ID AND game_key=@GAME_KEY";

        try
        {
            using (cmd = new SqlCommand(sql_check_key, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;
                cmd.Parameters.Add("@GAME_KEY", SqlDbType.VarChar, 50).Value = _game_key;
                using (dr = cmd.ExecuteReader())
                {
                    if (dr.Read())
                    {
                        is_valid = System.Convert.ToInt32(dr["id"]);
                    }
                }
            }
        }
        catch (Exception e)
        {
            log4.Error(e.ToString());
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
        return is_valid;
    }

    /**
     * 登入時更新
     */
    public void loginUpdate(string _login_id, string _game_key, string _ip_address, string nickname, string realname, string email)
    {
        string sql = "SELECT * FROM fishbowl_account WHERE login_id = @LOGIN_ID";

        try
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;

                using (dr = cmd.ExecuteReader())
                {
                    if (dr.Read())
                    {
                        // 有資料，更新
                        // 2009.10.20 修正可更新realname & nickname欄位
                        int account_id = System.Convert.ToInt32(dr["id"]);
                        dr.Close();

                        sql = "UPDATE fishbowl_account SET game_key=@GAME_KEY, realname = @REAL_NAME, nickname = @NICK_NAME, last_modify=@LAST_MODIFY WHERE id=@ID";
                        cmd.CommandText = sql;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@GAME_KEY", SqlDbType.VarChar, 50).Value = _game_key;
                        cmd.Parameters.Add("@ID", SqlDbType.Int, 4).Value = account_id;
                        cmd.Parameters.Add("@LAST_MODIFY", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.Parameters.Add("@REAL_NAME", SqlDbType.NVarChar, 200).Value = realname;
                        cmd.Parameters.Add("@NICK_NAME", SqlDbType.NVarChar, 200).Value = nickname;

                        cmd.ExecuteNonQuery();

                        log(account_id, "Login");
                    }
                    else
                    {
                        // 無資料，寫入
                        dr.Close();
                        sql = "INSERT INTO fishbowl_account(login_id, game_key, create_time, realname, nickname, email) ";
                        sql += " VALUES(@LOGIN_ID, @GAME_KEY, @CREATE_TIME, @REAL_NAME, @NICK_NAME, @EMAIL)";
                        cmd = new SqlCommand(sql, conn);
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;
                        cmd.Parameters.Add("@GAME_KEY", SqlDbType.VarChar, 50).Value = _game_key;
                        cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.Parameters.Add("@REAL_NAME", SqlDbType.VarChar, 200).Value = realname;
                        cmd.Parameters.Add("@NICK_NAME", SqlDbType.VarChar, 200).Value = nickname;
                        cmd.Parameters.Add("@EMAIL", SqlDbType.VarChar, 200).Value = email;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "Select @@Identity";
                        int last_id = 0;
                        object o = cmd.ExecuteScalar();
                        if (!(o is DBNull))
                        {
                            last_id = Convert.ToInt32(o);
                        }

                        log(last_id, "NewGame");

                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }
        catch (Exception e)
        {
            log4.Error(e.ToString());
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    /**
     * 使用者點擊
     */
    public void update_user_know(int fish_id)
    {
        sql = "UPDATE fishbowl_pets SET user_know = 1 WHERE id = @FISH_ID";

        try
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = fish_id;
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception e)
        {
            log4.Error(e.ToString());
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }

    /**
     * 遊戲重設
     */
    public void reset_data(int account_id)
    {
        // 抓出所有未長大的魚而且也還活著的魚(status = 1)，而且使用者未點擊過的(user_know = 0)
        // 並把狀態設定為(status = 3, user_know = 1);
        sql = "UPDATE fishbowl_pets SET status = 3, user_know = 1 WHERE status = 1 AND user_know = 0 AND account_id = @ACCOUNT_ID";

        try
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                cmd.ExecuteNonQuery();
            }

            log(account_id, "Reset");
        }
        catch (Exception e)
        {
            log4.Error(e.ToString());
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                conn.Close();
            }
        }
    }


    /**
     * LOGGER
     */
    private void log(int account_id, string action)
    {
        sql = "INSERT INTO fishbowl_account_log(account_id, action, ip_address, create_time) VALUES(@ACCOUNT_ID, @ACTION, @IPADDRESS, @CREATE_TIME)";

        try
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                cmd.Parameters.Add("@ACTION", SqlDbType.VarChar, 20).Value = action;
                //cmd.Parameters.Add("@IPADDRESS", SqlDbType.VarChar, 15).Value = System.Web.HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
                cmd.Parameters.Add("@IPADDRESS", SqlDbType.VarChar, 15).Value = FishBowlUtil.get_ip();
                cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                cmd.ExecuteNonQuery();
            }
        }
        catch (Exception e)
        {
            log4.Error(e.ToString());
        }
        finally
        {
            if (conn.State == ConnectionState.Open)
            {
                //conn.Close();
            }
        }
    }

}
