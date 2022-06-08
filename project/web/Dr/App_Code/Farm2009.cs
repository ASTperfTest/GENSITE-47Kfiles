using System;
using System.Collections;
using System.Collections.Generic;
using System.Web;
using System.Data;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.IO;
using log4net;

public class Farm2009
{
    private SqlCommand cmd;
    private SqlDataReader dr;
    private string _login_id = "";
    private string _game_key = "";
    private ILog log4 = LogManager.GetLogger("Farm2009");
	
    public Farm2009(string login_id, string game_key)
    {
        log4net.Config.XmlConfigurator.Configure();

        if (login_id != null)
        {
            _login_id = login_id.Trim();
        }

        if (game_key != null)
        {
            _game_key = game_key.Trim();
        }
    }

    /**
     * 取得account_id
     */
    public int account_id
    {
        get
        {
            int id = 0;

            string sql = "SELECT id FROM farm2009_account WHERE login_id = @LOGIN_ID";

            using (SqlConnection conn = FarmDB.createConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;
                    try
                    {
                        conn.Open();
                        object o = cmd.ExecuteScalar();
                        if (!(o is DBNull))
                        {
                            id = Convert.ToInt32(o);
                        }
                    }
                    catch (Exception ex)
                    {
                        log4.Error(ex.ToString());
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }
            }
            
            return id;
        }
    }

    /**
     * 取得該帳號目前最高輪迴數
     */
    private int current_round
    {
        get
        {
            return Convert.ToInt32(game_status["current_round"]);
        }
    }

    /**
     * 取得該帳號目前最高關卡數
     */
    private int current_level
    {
        get
        {
            return Convert.ToInt32(game_status["current_level"]);
        }
    }

    /**
     * 目前遊戲進度
     */
    private Hashtable game_status
    {
        get
        {
            Hashtable current = new Hashtable();
            current.Add("current_round", 1);
            current.Add("current_level", 0);

            string sql = "SELECT TOP 1 current_round, current_level FROM farm2009_question_data WHERE account_id = @ACCOUNT_ID ORDER BY current_round DESC, current_level DESC";

            using (SqlConnection conn = FarmDB.createConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = account_id;

                    try
                    {
                        conn.Open();
                        using (SqlDataReader dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                current["current_round"] = Convert.ToInt32(dr["current_round"]);
                                current["current_level"] = Convert.ToInt32(dr["current_level"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log4.Error(ex.ToString());
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }
            }

            return current;
        }
    }

    /**
     * 檢查是否為有效之Game key
     */
    public int isValidKey
    {
        get
        {
            int is_valid = 0;

            string sql_check_key = "SELECT * FROM farm2009_account WHERE login_id=@LOGIN_ID AND game_key=@GAME_KEY";

            using (SqlConnection conn = FarmDB.createConnection())
            {
                using (cmd = new SqlCommand(sql_check_key, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;
                    cmd.Parameters.Add("@GAME_KEY", SqlDbType.VarChar, 50).Value = _game_key;

                    try
                    {
                        conn.Open();
                        using (dr = cmd.ExecuteReader())
                        {
                            if (dr.Read())
                            {
                                is_valid = System.Convert.ToInt32(dr["id"]);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log4.Error(ex.ToString());
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }
            }

            return is_valid;
        }
    }

    /**
     * 登入時更新Game key
     */
    public void loginUpdate(string nickname, string realname, string email)
    {
        string sql = "SELECT id FROM farm2009_account WHERE login_id = @LOGIN_ID";

        using (SqlConnection conn = FarmDB.createConnection())
        {
            using (SqlCommand cmd_update = new SqlCommand(sql, conn))
            {
                cmd_update.Parameters.Clear();
                cmd_update.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;

                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd_update.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            // 有資料，更新
                            int account_id = Convert.ToInt32(dr["id"]);
                            dr.Close();

                            using (SqlCommand cmd_login = new SqlCommand())
                            {
                                sql = "UPDATE farm2009_account SET game_key=@GAME_KEY, realname=@REAL_NAME, nickname=@NICK_NAME, last_modify=@LAST_MODIFY WHERE id=@ID";

                                cmd_login.CommandText = sql;
                                cmd_login.Connection = conn;
                                cmd_login.Parameters.Clear();
                                cmd_login.Parameters.Add("@GAME_KEY", SqlDbType.VarChar, 50).Value = _game_key;
                                cmd_login.Parameters.Add("@REAL_NAME", SqlDbType.NVarChar, 200).Value = realname;
                                cmd_login.Parameters.Add("@NICK_NAME", SqlDbType.NVarChar, 200).Value = nickname;
                                cmd_login.Parameters.Add("@LAST_MODIFY", SqlDbType.DateTime).Value = DateTime.Now;
                                cmd_login.Parameters.Add("@ID", SqlDbType.Int).Value = account_id;

                                cmd_login.ExecuteNonQuery();
                            }
                            
                            log(account_id, "Login");
                        }
                        else
                        {
                            // 無資料，寫入
                            dr.Close();
                            sql = "INSERT INTO farm2009_account(login_id, game_key, create_time, realname, nickname, email) ";
                            sql += " VALUES(@LOGIN_ID, @GAME_KEY, @CREATE_TIME, @REAL_NAME, @NICK_NAME, @EMAIL)";

                            using (SqlCommand cmd_insert = new SqlCommand(sql, conn))
                            {
                                cmd_insert.Parameters.Clear();
                                cmd_insert.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;
                                cmd_insert.Parameters.Add("@GAME_KEY", SqlDbType.VarChar, 50).Value = _game_key;
                                cmd_insert.Parameters.Add("@REAL_NAME", SqlDbType.VarChar, 200).Value = realname;
                                cmd_insert.Parameters.Add("@NICK_NAME", SqlDbType.VarChar, 200).Value = nickname;
                                cmd_insert.Parameters.Add("@EMAIL", SqlDbType.VarChar, 200).Value = email;
                                cmd_insert.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                                cmd_insert.ExecuteNonQuery();

                                cmd_insert.CommandText = "Select @@Identity";
                                int last_id = 0;
                                object o = cmd_insert.ExecuteScalar();
                                if (!(o is DBNull))
                                {
                                    last_id = Convert.ToInt32(o);
                                }

                                log(last_id, "NewGame");

                                cmd_insert.ExecuteNonQuery();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                }
            }
        }
    }

    /**
     * LOGGER
     */
    private void log(int account_id, string action)
    {
        string sql = "INSERT INTO farm2009_account_log(account_id, action, ip_address, create_time) ";
        sql += "VALUES(@ACCOUNT_ID, @ACTION, @IPADDRESS, @CREATE_TIME)";

        using (SqlConnection conn = FarmDB.createConnection())
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = account_id;
                cmd.Parameters.Add("@ACTION", SqlDbType.VarChar, 30).Value = action;
                cmd.Parameters.Add("@IPADDRESS", SqlDbType.VarChar, 15).Value = ip_address;
                cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                }
            }
        }
    }

    /**
     * 驗證是否符合編碼
     */
    public bool isMatchedKey(string _key)
    {
        return (_key == MD5(_login_id + Farm2009Config.SALT + _game_key));
    }

    /**
     * MD5編碼
     */
    private string MD5(string Value)
    {
        MD5CryptoServiceProvider x = new MD5CryptoServiceProvider();
        byte[] data = Encoding.ASCII.GetBytes(Value);
        data = x.ComputeHash(data);
        string ret = "";
        for (int i = 0; i < data.Length; i++)
        {
            ret += data[i].ToString("x2").ToLower();
        }
        return ret;
    }

    /**
     * 取題目
     */
    public string getQuestion(int level, int count)
    {
        string xml = "";
        
        string sql = "SELECT TOP (@COUNT) * FROM farm2009_questions WHERE qrank=@LEVEL and disable=0 ORDER BY NEWID()";

        using (SqlConnection conn = FarmDB.createConnection())
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@COUNT", SqlDbType.TinyInt).Value = count;
                cmd.Parameters.Add("@LEVEL", SqlDbType.TinyInt).Value = level;

                try
                {
                    conn.Open();
                    using (dr = cmd.ExecuteReader())
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                        sb.Append("<data>");
                        while (dr.Read())
                        {
                            sb.Append("<qa>");
                            sb.Append("<qrank>" + dr["qrank"].ToString() + "</qrank>");

                            sb.Append("<qcontent>" + dr["qcontent"].ToString() + "</qcontent>");
                            sb.Append("<a1>" + dr["a1"].ToString() + "</a1>");
                            sb.Append("<a2>" + dr["a2"].ToString() + "</a2>");
                            sb.Append("<a3>" + dr["a3"].ToString() + "</a3>");
                            sb.Append("<correct>" + dr["correct"].ToString() + "</correct>");
                            sb.Append("<qid>" + dr["id"].ToString() + "</qid>");
                            sb.Append("</qa>");
                        }
                        sb.Append("</data>");
                        xml = sb.ToString();
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                }
            }
        }

        return xml;
    }

    /**
     * 寫入遊戲資料
     */
    public int updateGameData(string _current_round, string _current_level, string _current_score, string _current_timer)
    {
        // 2010.2.24 修正
        // 檢查level值是否正常
        if (isValidLevel(_current_level))
        {
            int last_id = 0;
            int the_account_id = account_id;

            using (SqlConnection conn = FarmDB.createConnection())
            {
                // 先檢查是否有重複資料
                string sql_check_duplicate = "SELECT id FROM farm2009_question_data WHERE current_round = @CURRENT_ROUND AND current_level = @CURRENT_LEVEL AND account_id = @ACCOUNT_ID";

                try
                {
                    conn.Open();
                    using (SqlCommand cmd_chk = new SqlCommand())
                    {
                        cmd_chk.Connection = conn;
                        cmd_chk.CommandText = sql_check_duplicate;
                        cmd_chk.Parameters.Clear();
                        cmd_chk.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = the_account_id;
                        cmd_chk.Parameters.Add("@CURRENT_LEVEL", SqlDbType.TinyInt).Value = Convert.ToInt32(_current_level);
                        cmd_chk.Parameters.Add("@CURRENT_ROUND", SqlDbType.Int).Value = Convert.ToInt32(_current_round);

                        string sql = "";

                        object data_obj = cmd_chk.ExecuteScalar();
                        if (!(data_obj is DBNull))
                        {
                            last_id = Convert.ToInt32(data_obj);
                        }

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {
                            cmd.Parameters.Clear();
                            if (last_id > 0)
                            {
                                // update
                                sql = "UPDATE farm2009_question_data SET current_score = @CURRENT_SCORE, current_timer = @CURRENT_TIMER ";
                                sql += "WHERE account_id = @ACCOUNT_ID AND id = @DATA_ID";

                                cmd.CommandText = sql;
                                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = the_account_id;
                                cmd.Parameters.Add("@CURRENT_SCORE", SqlDbType.Int).Value = Convert.ToInt32(_current_score);
                                cmd.Parameters.Add("@CURRENT_TIMER", SqlDbType.BigInt).Value = Convert.ToInt32(_current_timer);
                                cmd.Parameters.Add("@DATA_ID", SqlDbType.Int).Value = last_id;

                                cmd.ExecuteNonQuery();
                            }
                            else
                            {
                                // insert
                                sql = "INSERT INTO farm2009_question_data(account_id, current_level, current_score, current_timer, create_time, current_round) ";
                                sql += "VALUES(@ACCOUNT_ID, @CURRENT_LEVEL, @CURRENT_SCORE, @CURRENT_TIMER, @CREATE_TIME, @CURRENT_ROUND)";

                                cmd.CommandText = sql;
                                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = the_account_id;
                                cmd.Parameters.Add("@CURRENT_LEVEL", SqlDbType.TinyInt).Value = Convert.ToInt32(_current_level);
                                cmd.Parameters.Add("@CURRENT_SCORE", SqlDbType.Int).Value = Convert.ToInt32(_current_score);
                                cmd.Parameters.Add("@CURRENT_TIMER", SqlDbType.BigInt).Value = Convert.ToInt32(_current_timer);
                                cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                                cmd.Parameters.Add("@CURRENT_ROUND", SqlDbType.Int).Value = Convert.ToInt32(_current_round);

                                cmd.ExecuteNonQuery();

                                cmd.CommandText = "Select @@Identity";
                                object o = cmd.ExecuteScalar();
                                if (!(o is DBNull))
                                {
                                    last_id = Convert.ToInt32(o);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                }
            }

            // 紀錄
            log(the_account_id, "save game data " + last_id);

            return last_id;
        }
        else
        {
            return 0;
        }
    }

    /**
     * 寫入答題紀錄
     */
    public void updateQuestionLog(int data_id, int question_order, int question_id, int is_correct)
    {
        // 2010.2.23 修正
        // 先檢查是否有該答題ID
        // 若該data_id無效則不寫入
        if (isValidDataId(data_id))
        {
            string sql = "INSERT INTO farm2009_question_log(data_id, question_id, question_order, is_correct, create_time) ";
            sql += "VALUES(@DATA_ID, @QUESTION_ID, @QUESTION_ORDER, @IS_CORRECT, @CREATE_TIME)";

            using (SqlConnection conn = FarmDB.createConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@DATA_ID", SqlDbType.Int).Value = data_id;
                    cmd.Parameters.Add("@QUESTION_ID", SqlDbType.Int).Value = question_id;
                    cmd.Parameters.Add("@QUESTION_ORDER", SqlDbType.TinyInt).Value = question_order;
                    cmd.Parameters.Add("@IS_CORRECT", SqlDbType.TinyInt).Value = is_correct;
                    cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;

                    try
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        log4.Error(ex.ToString());
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }
            }
        }
    }

    /**
     * 取得遊戲資料
     */
    public string getGameData()
    {
        string xml = "";
        SqlCommand cmd;
        
        int the_current_round = current_round;
        int the_current_level = current_level;
        int the_account_id = account_id;

        string sql_current = "SELECT SUM(current_score) as total_score, SUM(current_timer) as total_timer FROM farm2009_question_data ";
        sql_current += "WHERE account_id = @ACCOUNT_ID AND current_round = @CURRENT_ROUND";

        using (SqlConnection conn = FarmDB.createConnection())
        {
            using (cmd = new SqlCommand())
            {
                cmd.Connection = conn;
                cmd.CommandText = sql_current;
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = the_account_id;
                cmd.Parameters.Add("@CURRENT_ROUND", SqlDbType.Int).Value = the_current_round;

                StringBuilder sb = new StringBuilder();
                sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sb.Append("<data>");
                sb.Append("<login_id>" + _login_id + "</login_id>");
                sb.Append("<game_key>" + _game_key + "</game_key>");

                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            sb.Append("<nowlevel>" + the_current_level.ToString() + "</nowlevel>");
                            if (dr["total_score"] is DBNull)
                            {
                                sb.Append("<nowScore>0</nowScore>");
                            }
                            else
                            {
                                sb.Append("<nowScore>" + Convert.ToInt32(dr["total_score"]) + "</nowScore>");
                            }

                            if (dr["total_timer"] is DBNull)
                            {
                                sb.Append("<nowtime>0</nowtime>");
                            }
                            else
                            {
                                sb.Append("<nowtime>" + Convert.ToInt32(dr["total_timer"]) + "</nowtime>");
                            }
                        }
                        sb.Append("<nowround>" + the_current_round.ToString() + "</nowround>");
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                }
                finally
                {
                    if (conn.State == ConnectionState.Open)
                    {
                        conn.Close();
                    }
                }

                // ============================================================

                // 歷史紀錄
                // 如果第N輪已結束(current_level = 3)
                // 抓取到目前最高的分數、最低的時間及最佳關數
                // 若未完成，則僅抓取N-1輪之資料

                string sql_best_score = "";

                Hashtable current_game_status = game_status;

                if (Convert.ToInt32(current_game_status["current_level"]) == 3)
                {
                    // 玩到第3關，該輪已結束
                    // 已至少過一大關，最高關卡一定為3
                    sb.Append("<bestlevel>3</bestlevel>");

                    // 抓取到目前最高的分數
                    sql_best_score = "SELECT TOP 1 current_round, SUM(current_score) AS total_score, SUM(current_timer) as total_timer ";
                    sql_best_score += "FROM farm2009_question_data ";
                    sql_best_score += "WHERE account_id = @ACCOUNT_ID GROUP BY current_round ORDER BY total_score DESC";

                    try
                    {
                        conn.Open();
                        using (SqlCommand cmd_score = new SqlCommand(sql_best_score, conn))
                        {
                            cmd_score.Parameters.Clear();
                            cmd_score.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = the_account_id;

                            using (dr = cmd_score.ExecuteReader())
                            {
                                if (dr.Read())
                                {
                                    sb.Append("<bestScore>" + dr["total_score"].ToString() + "</bestScore>");
                                    sb.Append("<besttime>" + dr["total_timer"].ToString() + "</besttime>");
                                }
                                else
                                {
                                    sb.Append("<bestScore>0</bestScore>");
                                    sb.Append("<besttime>0</besttime>");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log4.Error(ex.ToString());
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }
                else
                {
                    if (Convert.ToInt32(current_game_status["current_round"]) > 1)
                    {
                        // 若大於第一輪，最高關卡一定是3
                        sb.Append("<bestlevel>3</bestlevel>");
                    }
                    else
                    {
                        // 否則為目前最高關卡
                        sb.Append("<bestlevel>" + current_game_status["current_level"].ToString() + "</bestlevel>");
                    }

                    // 未完成，則僅抓取N-1輪之資料
                    // 抓取到目前最高的分數
                    sql_best_score = "SELECT TOP 1 current_round, SUM(current_score) AS total_score, SUM(current_timer) as total_timer ";
                    sql_best_score += "FROM farm2009_question_data ";
                    sql_best_score += "WHERE account_id = @ACCOUNT_ID AND current_round <> @CURRENT_ROUND GROUP BY current_round ORDER BY total_score DESC";

                    try
                    {   
                        conn.Open();
                        using (SqlCommand cmd_score = new SqlCommand(sql_best_score, conn))
                        {
                            cmd_score.Parameters.Clear();
                            cmd_score.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = the_account_id;
                            cmd_score.Parameters.Add("@CURRENT_ROUND", SqlDbType.Int).Value = current_game_status["current_round"];

                            using (dr = cmd_score.ExecuteReader())
                            {
                                if (dr.Read())
                                {
                                    sb.Append("<bestScore>" + dr["total_score"].ToString() + "</bestScore>");
                                    sb.Append("<besttime>" + dr["total_timer"].ToString() + "</besttime>");
                                }
                                else
                                {
                                    sb.Append("<bestScore>0</bestScore>");
                                    sb.Append("<besttime>0</besttime>");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log4.Error(ex.ToString());
                    }
                    finally
                    {
                        if (conn.State == ConnectionState.Open)
                        {
                            conn.Close();
                        }
                    }
                }

                using (SqlCommand cmd_nickname = new SqlCommand())
                {
                    string sql = "SELECT nickname, realname FROM farm2009_account WHERE id = @ID";
                    cmd_nickname.Connection = conn;
                    cmd_nickname.CommandText = sql;
                    cmd_nickname.Parameters.Clear();
                    cmd_nickname.Parameters.Add("@ID", SqlDbType.Int, 4).Value = the_account_id;
                    conn.Open();
                    using (dr = cmd_nickname.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            string nickname = dr["nickname"].ToString();
                            string realname = dr["realname"].ToString();

                            sb.Append("<nick_name>" + replaceRealName(nickname, realname) + "</nick_name>");
                        }
                        else
                        {
                            sb.Append("<nick_name></nick_name>");
                        }
                    }
                    conn.Close();
                }

                sb.Append("</data>");
                xml = sb.ToString();
            }
        }

        return xml;
    }

    /**
     * 取得IP位置
     */
    private string ip_address
    {
        get
        {
            return System.Web.HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
        }
    }

    /**
     * 把真實姓名馬賽克
     */
    private string replaceRealName(string nickname, string realname)
    {
        if (nickname != null && nickname.Trim() != "")
        {
            return nickname.Trim();
        }
        else
        {
            return (realname.Substring(0, 1) + "＊" + realname.Substring(realname.Trim().Length - 1, 1));
        }
    }

    /**
     * 取得前N名的榜行
     */
    public string getTopScore(int count)
    {
        StringBuilder sb = new StringBuilder();

        int the_account_id = account_id;

        sb.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sb.Append("<data>");

        using (SqlConnection conn = FarmDB.createConnection())
        {
            // 個人資料區
            sb.Append("<personal>");
            using (SqlCommand cmd_nickname = new SqlCommand())
            {
                string sql = "SELECT nickname, realname FROM farm2009_account WHERE id = @ID";
                cmd_nickname.Connection = conn;
                cmd_nickname.CommandText = sql;
                cmd_nickname.Parameters.Clear();
                cmd_nickname.Parameters.Add("@ID", SqlDbType.Int, 4).Value = the_account_id;
                conn.Open();
                using (dr = cmd_nickname.ExecuteReader())
                {
                    if (dr.Read())
                    {
                        string nickname = dr["nickname"].ToString();
                        string realname = dr["realname"].ToString();

                        sb.Append("<nick_name>" + replaceRealName(nickname, realname) + "</nick_name>");
                    }
                    else
                    {
                        sb.Append("<nick_name></nick_name>");
                    }
                }
                conn.Close();
            }

            sb.Append("<login_id>" +_login_id +"</login_id>");

            // 個人紀錄

            StringBuilder sqlPersonal = new StringBuilder();
            sqlPersonal.Append(" SELECT account.id, account.login_id, account.realname, account.nickname, game_data.top_level, game_data.top_score, game_data.top_timer");
            sqlPersonal.Append(" FROM farm2009_account AS account INNER JOIN ");
            sqlPersonal.Append(" (SELECT account_id, MAX(max_level) AS top_level, MAX(total_score) AS top_score, MIN(total_timer) AS top_timer ");
            sqlPersonal.Append(" FROM ( ");
            sqlPersonal.Append(" SELECT MIN(id) as game_data_id, account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer FROM farm2009_question_data GROUP BY account_id, current_round ");
            sqlPersonal.Append(" ) AS game_data_1 ");
            sqlPersonal.Append(" GROUP BY account_id ");
            sqlPersonal.Append(" ) AS game_data ON account.id = game_data.account_id AND id = @ACCOUNT_ID ");
            sqlPersonal.Append(" ORDER BY game_data.top_level DESC, game_data.top_score DESC, game_data.top_timer, id");

            using (SqlCommand cmd_personal = new SqlCommand(sqlPersonal.ToString(), conn))
            {
                cmd_personal.Parameters.Clear();
                cmd_personal.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = the_account_id;
                conn.Open();
                using (dr = cmd_personal.ExecuteReader())
                {
                    if (dr.Read())
                    {
                        sb.Append("<bestlevel>" + dr["top_level"].ToString() + "</bestlevel>");
                        sb.Append("<bestScore>" + dr["top_score"].ToString() + "</bestScore>");
                        sb.Append("<besttime>" + getGameTimer(the_account_id, Convert.ToInt32(dr["top_level"]), Convert.ToInt32(dr["top_score"])) + "</besttime>");
                    }
                    else
                    {
                        sb.Append("<bestlevel>0</bestlevel>");
                        sb.Append("<bestScore>0</bestScore>");
                        sb.Append("<besttime>0</besttime>");
                    }
                }
                conn.Close();
            }

            // 個人排行
            sb.Append("<ranking>" +getPersonalRank(the_account_id).ToString()  +"</ranking>");

            sb.Append("</personal>");

            // 其它人排名
            sb.Append("<all>");

            using(SqlCommand cmd_all = new SqlCommand())
            {
                StringBuilder sqlString = new StringBuilder();
                
                //sqlString.Append(" SELECT TOP (@COUNTER) account.id, account.login_id, account.realname, account.nickname, game_data.top_level, game_data.top_score, game_data.top_timer");
                //sqlString.Append(" FROM farm2009_account AS account INNER JOIN ");
                //sqlString.Append(" (SELECT account_id, MAX(max_level) AS top_level, MAX(total_score) AS top_score, MIN(total_timer) AS top_timer ");
                //sqlString.Append(" FROM ( ");
                //sqlString.Append(" SELECT account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer FROM farm2009_question_data GROUP BY account_id, current_round ");
                //sqlString.Append(" ) AS game_data_1 ");
                //sqlString.Append(" GROUP BY account_id ");
                //sqlString.Append(" ) AS game_data ON account.id = game_data.account_id ");
                //sqlString.Append(" ORDER BY game_data.top_level DESC, game_data.top_score DESC, game_data.top_timer, id");

                // 2009.11.16 fixed by eddie
                sqlString.Append(" SELECT TOP (@COUNTER) account.id, account.login_id, account.realname, account.nickname, game_data.top_level, game_data.top_score, game_data.top_timer ");
                sqlString.Append(" FROM farm2009_account account, ");
                sqlString.Append(" ( ");
                sqlString.Append(" SELECT account_id, MAX(max_level) AS top_level, MAX(total_score) AS top_score, MIN(best_timer) AS top_timer  ");
                sqlString.Append(" FROM ");
                sqlString.Append(" ( ");
                sqlString.Append(" SELECT account_id, MAX(q_data.current_level) AS max_level, SUM(q_data.current_score) AS total_score, SUM(q_data.current_timer) AS total_timer,  ");
                sqlString.Append(" ( ");
                sqlString.Append(" SELECT TOP 1 total_timer FROM  ");
                sqlString.Append(" ( ");
                sqlString.Append(" SELECT account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer ");
                sqlString.Append(" FROM farm2009_question_data  ");
                sqlString.Append(" GROUP BY account_id, current_round ");
                sqlString.Append(" ) AS timer_table ");
                sqlString.Append(" WHERE (timer_table.account_id = q_data.account_id) AND (timer_table.max_level = max_level) AND (timer_table.total_score = total_score) ");
                sqlString.Append(" ) AS best_timer ");
                sqlString.Append(" FROM farm2009_question_data q_data ");
                sqlString.Append(" GROUP BY q_data.account_id, q_data.current_round ");
                sqlString.Append(" ) AS score_list ");
                sqlString.Append(" GROUP BY account_id ");
                sqlString.Append(" ) AS game_data ");
                sqlString.Append(" WHERE account.id = game_data.account_id ");
                sqlString.Append(" ORDER BY game_data.top_level DESC, game_data.top_score DESC, game_data.top_timer ");

                cmd_all.Connection = conn;
                cmd_all.CommandText = sqlString.ToString();
                cmd_all.Parameters.Clear();
                cmd_all.Parameters.Add("@COUNTER", SqlDbType.Int).Value = count;

                conn.Open();
                using (dr = cmd_all.ExecuteReader())
                {
                    string nickname;
                    string realname;
                    while (dr.Read())
                    {
                        nickname = dr["nickname"].ToString();
                        realname = dr["realname"].ToString();
                        sb.Append("<list>");
                        sb.Append("<nick_name>" + replaceRealName(nickname, realname) + "</nick_name>");
                        sb.Append("<login_id>" +dr["login_id"].ToString() +"</login_id>");
                        sb.Append("<bestlevel>" +dr["top_level"].ToString() +"</bestlevel>");
                        sb.Append("<bestScore>" +dr["top_score"].ToString() +"</bestScore>");
                        //sb.Append("<besttime>" + getGameTimer(Convert.ToInt32(dr["id"]), Convert.ToInt32(dr["top_level"]), Convert.ToInt32(dr["top_score"])) +"</besttime>");
                        // 2009.11.16 fixed by eddie
                        //sb.Append("<besttime>" + SecsToMins(Convert.ToInt32(dr["top_timer"])).ToString() + "</besttime>");
                        // 2009.11.18 fixed by eddie
                        sb.Append("<besttime>" + dr["top_timer"].ToString() + "</besttime>");
                        sb.Append("</list>");
                    }
                }
                conn.Close();
            }

            sb.Append("</all>");
        }

        sb.Append("</data>");

        return sb.ToString();
    }

    private string getGameTimer(int account_id, int top_level, int top_score)
    {
        string top_timer = "";

        using (SqlConnection conn = FarmDB.createConnection())
        {
            using(SqlCommand cmd_all = new SqlCommand())
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT total_timer FROM (SELECT account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer");
                sb.Append(" FROM farm2009_question_data GROUP BY account_id, current_round) AS derivedtbl_1 WHERE (account_id = @ACCOUNT_ID) AND (max_level = @MAX_LEVEL) AND (total_score = @MAX_SCORE)");
                
                cmd_all.Connection = conn;
                cmd_all.CommandText = sb.ToString();
                cmd_all.Parameters.Clear();
                cmd_all.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                cmd_all.Parameters.Add("@MAX_LEVEL", SqlDbType.Int, 4).Value = top_level;
                cmd_all.Parameters.Add("@MAX_SCORE", SqlDbType.Int, 4).Value = top_score;

                conn.Open();
                using (SqlDataReader the_dr = cmd_all.ExecuteReader())
                {
                    if (the_dr.Read())
                    {
                        top_timer = the_dr["total_timer"].ToString();
                    }
                }
                conn.Close();
            }
        }

        return top_timer;
    }

    /**
     * echo
     */
    private void echo(string str)
    {
        System.Web.HttpContext.Current.Response.Write(str);
    }

    /**
     * 取得個人排行
     */
    private int getPersonalRank(int the_account_id)
    {
        int rank = 0;

        string sql = "DECLARE @RankingTable TABLE (ranking int, account_id int, top_level int, top_score int, top_timer int) ";
        sql += " INSERT INTO @RankingTable (ranking, account_id, top_level, top_score, top_timer) ";
        sql += " SELECT row_number() OVER(ORDER BY top_level DESC, top_score DESC, game_data.top_timer) as ranking, account.id, game_data.top_level, game_data.top_score, game_data.top_timer ";
        sql += " FROM farm2009_account AS account INNER JOIN ";
        sql += " (SELECT account_id, MAX(max_level) AS top_level, MAX(total_score) AS top_score, MIN(total_timer) AS top_timer ";
        sql += " FROM (SELECT account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer ";
        sql += " FROM farm2009_question_data ";
        sql += " GROUP BY account_id, current_round) AS game_data_1 ";
        sql += " GROUP BY account_id) AS game_data ON account.id = game_data.account_id ";
        sql += " ORDER BY game_data.top_level DESC, game_data.top_score DESC, game_data.top_timer ";
        sql += " select * from @RankingTable WHERE account_id = @ACCOUNT_ID";

        try
        {
            using (SqlConnection conn = FarmDB.createConnection())
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = the_account_id;

                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        rank = Convert.ToInt32(o);
                    }
                }
                conn.Close();
            }
        }
        catch (Exception ex)
        {
            log4.Error(ex.ToString());
        }
        return rank;
    }

    // 2009.11.16 added by eddie
    private string SecsToMins(int top_timer)
    {
        int mins = top_timer / 60;
        int secs = top_timer % 60;
        return mins + "分" + secs + "秒";
    }

    /**
     * 檢查是否為有效之data_id
     * 2010.2.23 修正
     */
    private bool isValidDataId(int data_id)
    {
        bool is_valid = false;

        using (SqlConnection conn = FarmDB.createConnection())
        {
            using (SqlCommand cmd_all = new SqlCommand())
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("SELECT id FROM farm2009_question_data WHERE id=@DATA_ID");

                cmd_all.Connection = conn;
                cmd_all.CommandText = sb.ToString();
                cmd_all.Parameters.Clear();
                cmd_all.Parameters.Add("@DATA_ID", SqlDbType.Int, 4).Value = data_id;

                conn.Open();
                using (SqlDataReader the_dr = cmd_all.ExecuteReader())
                {
                    if (the_dr.Read())
                    {
                        is_valid = true;
                    }
                }
                conn.Close();
            }
        }

        return is_valid;
    }

    /**
     * 是否為正確的nowlevel
     * 2010.2.24 修正
     */
    private bool isValidLevel(string level)
    {
        // level只有1、2、3這三種結果
        return (level == "1" || level == "2" || level == "3");
    }
}
