using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Text;
using log4net;

public partial class game_data : System.Web.UI.Page
{
    ILog log4 = LogManager.GetLogger("FishBowl");
    protected void Page_Load(object sender, EventArgs e)
    {
        //Thread.Sleep(10000);
        log4net.Config.XmlConfigurator.Configure();
        string _login_id = Request.Form["login_id"].ToString();
        string _game_key = Request.Form["game_key"].ToString();
        string _key = Request.QueryString["key"].ToString();
        StringBuilder sb;

        if (!FishBowlUtil.isMatchedKey(_login_id, _game_key, _key))
        {
            Response.Write("<script>alert('請重新登入!');location.href='/';</script>");
            Response.End();
            return;
        }

        if (_login_id.Length <=0 || _game_key.Length <= 0)
        {
            Response.Write("<script>alert('請重新登入!');location.href='/';</script>");
            Response.End();
            return;
        }
        else
        {
            // 檢查該key是否正確
            SqlConnection conn = SQLDB.createConnection();

            try
            {

                string sql_check_key = "SELECT * FROM fishbowl_account WHERE login_id=@LOGIN_ID AND game_key=@GAME_KEY";

                SqlCommand cmd = new SqlCommand(sql_check_key, conn);
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = _login_id;
                cmd.Parameters.Add("@GAME_KEY", SqlDbType.VarChar, 50).Value = _game_key;

                SqlDataReader dr = cmd.ExecuteReader();

                if (!dr.Read())
                {
                    // Game key錯誤
                    dr.Close();
                }
                else
                {
                    // KEY正確，取出資料
                    int account_id = System.Convert.ToInt32(dr["id"]);
                    dr.Close();

                    // ===========環境變數================
                    int intMoney = 50;
                    int intStars = 0;
                    int intScore = 0;
                    double intWaterstatus = 7;
                    long intPlayTime = 0;

                    // ===========養殖物狀態================
                    // fish type 1 = Anemone(A)
                    // fish type 2 = Bream(B)
                    // fish type 3 = Clownfish(C)
                    // fish type 4 = Hippocampus(H)

                    int anemone_id = 0;
                    long anemoneBirthDay = 0;
                    double anemoneFullStatus = 4;
                    double anemoneHealthStatus = 10;
                    double anemoneScore = 0;
                    double anemoneScoreHistory = 0;
                    int anemoneStar = 0;
                    int anemoneStatus = 0;

                    int bream_id = 0;
                    long breamBirthDay = 0;
                    double breamFullStatus = 4;
                    double breamHealthStatus = 10;
                    double breamScore = 0;
                    double breamScoreHistory = 0;
                    int breamStar = 0;
                    int breamStatus = 0;

                    int clownfish_id = 0;
                    long clownfishBirthDay = 0;
                    double clownfishFullStatus = 4;
                    double clownfishHealthStatus = 10;
                    double clownfishScore = 0;
                    double clownfishScoreHistory = 0;
                    int clownfishStar = 0;
                    int clownfishStatus = 0;

                    int hippocampus_id = 0;
                    long hippocampusBirthDay = 0;
                    double hippocampusFullStatus = 4;
                    double hippocampusHealthStatus = 10;
                    double hippocampusScore = 0;
                    double hippocampusScoreHistory = 0;
                    int hippocampusStar = 0;
                    int hippocampusStatus = 0;

                    // 檢查是否有環境資料，若無則寫入一筆預設值
                    string sql_get_environment = "SELECT * FROM fishbowl_environment WHERE account_id=@ACCOUNT_ID";

                    cmd = new SqlCommand(sql_get_environment, conn);
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;

                    dr = cmd.ExecuteReader();

                    if (dr.Read())
                    {
                        // 有環境資料
                        intMoney = System.Convert.ToInt32(dr["money"]) + System.Convert.ToInt32(dr["money_add"]);
                        intWaterstatus = System.Convert.ToDouble(dr["water_status"]) + System.Convert.ToDouble(dr["water_status_minus"]);
                        intPlayTime = (dr["create_time"] == DBNull.Value) ? 0 : FishBowlUtil.getTimeDelta(System.Convert.ToDateTime(dr["create_time"]));

                        int environment_id = System.Convert.ToInt32(dr["id"]);
                        dr.Close();

                        // update water_status欄位
                        string sql_update_water_status = "UPDATE fishbowl_environment SET water_status = @WATER_STATUS, water_status_minus = 0 WHERE id=@ENVIRONMENT_ID";
                        cmd.CommandText = sql_update_water_status;

                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@WATER_STATUS", SqlDbType.Float).Value = intWaterstatus;
                        cmd.Parameters.Add("@ENVIRONMENT_ID", SqlDbType.Int, 4).Value = environment_id;
                        cmd.ExecuteNonQuery();

                        // 更新money_add欄位
                        string sql_update_money_add = "UPDATE fishbowl_environment SET money_add = 0 WHERE id = @ENVIRONMENT_ID";
                        cmd.CommandText = sql_update_money_add;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@ENVIRONMENT_ID", SqlDbType.Int, 4).Value = environment_id;
                        cmd.ExecuteNonQuery();

                        // 取得環境分數(水質)
                        /*string sql_get_environment_score = "SELECT SUM(score) FROM fishbowl_environment_score WHERE environment_id = @ENVIRONMENT_ID";
                        cmd.CommandText = sql_get_environment_score;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@ENVIRONMENT_ID", SqlDbType.Int, 4).Value = environment_id;
                        object s = cmd.ExecuteScalar();
                        if (!(s is DBNull))
                        {
                            intScore = System.Convert.ToInt32(s);
                        }*/
                    }
                    else
                    {
                        // 無環境資料，新增一筆
                        dr.Close();

                        intMoney = System.Convert.ToInt32(FishBowlUtil.getConfig("starting_money"));

                        string sql_add_environment = "INSERT INTO fishbowl_environment(account_id, create_time, money) VALUES(@ACCOUNT_ID, @CREATE_TIME, @MONEY)";
                        cmd.CommandText = sql_add_environment;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                        cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.Parameters.Add("@MONEY", SqlDbType.Int, 4).Value = intMoney;
                        cmd.ExecuteNonQuery();

                        cmd.CommandText = "Select @@Identity";
                        int last_environment_id = 0;
                        object o = cmd.ExecuteScalar();
                        if (!(o is DBNull))
                        {
                            last_environment_id = Convert.ToInt32(o);
                        }

                        /*string sql_add_environment_score = "INSERT INTO fishbowl_environment_score(environment_id, create_time) VALUES(@ENVIRONMENT_ID, @CREATE_TIME)";
                        cmd.CommandText = sql_add_environment_score;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@ENVIRONMENT_ID", SqlDbType.Int, 4).Value = last_environment_id;
                        cmd.Parameters.Add("@CREATE_TIME", SqlDbType.DateTime).Value = DateTime.Now;
                        cmd.ExecuteNonQuery();*/
                    }

                    // 取得最新的一筆資料，且user_know = 0
                    string sql_check_fishes = "SELECT TOP 1 * FROM fishbowl_pets WHERE fish_type=@FISH_TYPE AND user_know = 0 AND account_id=@ACCOUNT_ID AND status > 0 ORDER BY create_time DESC";
                    string sql_update_fishes = "UPDATE fishbowl_pets SET health_status = @NEW_HEALTH_STATUS, health_status_minus = 0, full_status = @NEW_FULL_STATUS, full_status_minus = 0 WHERE id = @FISH_ID";
                    // ====== 取得養殖品 1 =======
                    cmd.CommandText = sql_check_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 1;
                    dr = cmd.ExecuteReader();

                    if (dr.Read())
                    {
                        anemone_id = System.Convert.ToInt32(dr["id"]);
                        anemoneBirthDay = (dr["create_time"] == DBNull.Value) ? 0 : FishBowlUtil.getTimeDelta(System.Convert.ToDateTime(dr["create_time"]));
                        anemoneFullStatus = System.Convert.ToDouble(dr["full_status"]) + System.Convert.ToDouble(dr["full_status_minus"]);
                        anemoneHealthStatus = System.Convert.ToDouble(dr["health_status"]) + System.Convert.ToDouble(dr["health_status_minus"]);
                        anemoneStatus = System.Convert.ToInt32(dr["status"]);
                    }
                    dr.Close();

                    // 更新欄位
                    cmd.CommandText = sql_update_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@NEW_HEALTH_STATUS", SqlDbType.Float, 4).Value = anemoneHealthStatus;
                    cmd.Parameters.Add("@NEW_FULL_STATUS", SqlDbType.Float, 4).Value = anemoneFullStatus;
                    cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = anemone_id;
                    cmd.ExecuteNonQuery();

                    // ====== 取得養殖品 2 =======
                    cmd.CommandText = sql_check_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 2;
                    dr = cmd.ExecuteReader();

                    if (dr.Read())
                    {
                        bream_id = System.Convert.ToInt32(dr["id"]);
                        breamBirthDay = (dr["create_time"] == DBNull.Value) ? 0 : FishBowlUtil.getTimeDelta(System.Convert.ToDateTime(dr["create_time"]));
                        breamFullStatus = System.Convert.ToDouble(dr["full_status"]) + System.Convert.ToDouble(dr["full_status_minus"]);
                        breamHealthStatus = System.Convert.ToDouble(dr["health_status"]) + System.Convert.ToDouble(dr["health_status_minus"]);
                        breamStar = System.Convert.ToInt32(dr["stars"]);
                        breamStatus = System.Convert.ToInt32(dr["status"]);
                    }
                    dr.Close();

                    // 更新欄位
                    cmd.CommandText = sql_update_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@NEW_HEALTH_STATUS", SqlDbType.Float, 4).Value = breamHealthStatus;
                    cmd.Parameters.Add("@NEW_FULL_STATUS", SqlDbType.Float, 4).Value = breamFullStatus;
                    cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = bream_id;
                    cmd.ExecuteNonQuery();

                    // ====== 取得養殖品 3 =======
                    cmd.CommandText = sql_check_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 3;
                    dr = cmd.ExecuteReader();

                    if (dr.Read())
                    {
                        clownfish_id = System.Convert.ToInt32(dr["id"]);
                        clownfishBirthDay = (dr["create_time"] == DBNull.Value) ? 0 : FishBowlUtil.getTimeDelta(System.Convert.ToDateTime(dr["create_time"]));
                        clownfishFullStatus = System.Convert.ToDouble(dr["full_status"]) + System.Convert.ToDouble(dr["full_status_minus"]);
                        clownfishHealthStatus = System.Convert.ToDouble(dr["health_status"]) + System.Convert.ToDouble(dr["health_status_minus"]);
                        clownfishStar = System.Convert.ToInt32(dr["stars"]);
                        clownfishStatus = System.Convert.ToInt32(dr["status"]);
                    }
                    dr.Close();

                    // 更新欄位
                    cmd.CommandText = sql_update_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@NEW_HEALTH_STATUS", SqlDbType.Float, 4).Value = clownfishHealthStatus;
                    cmd.Parameters.Add("@NEW_FULL_STATUS", SqlDbType.Float, 4).Value = clownfishFullStatus;
                    cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = clownfish_id;
                    cmd.ExecuteNonQuery();

                    // ====== 取得養殖品 4 =======
                    cmd.CommandText = sql_check_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 4;
                    dr = cmd.ExecuteReader();

                    if (dr.Read())
                    {
                        hippocampus_id = System.Convert.ToInt32(dr["id"]);
                        hippocampusBirthDay = (dr["create_time"] == DBNull.Value) ? 0 : FishBowlUtil.getTimeDelta(System.Convert.ToDateTime(dr["create_time"]));
                        hippocampusFullStatus = System.Convert.ToDouble(dr["full_status"]) + System.Convert.ToDouble(dr["full_status_minus"]);
                        hippocampusHealthStatus = System.Convert.ToDouble(dr["health_status"]) + System.Convert.ToDouble(dr["health_status_minus"]);
                        hippocampusStar = System.Convert.ToInt32(dr["stars"]);
                        hippocampusStatus = System.Convert.ToInt32(dr["status"]);
                    }
                    dr.Close();

                    // 更新欄位
                    cmd.CommandText = sql_update_fishes;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@NEW_HEALTH_STATUS", SqlDbType.Float, 4).Value = hippocampusHealthStatus;
                    cmd.Parameters.Add("@NEW_FULL_STATUS", SqlDbType.Float, 4).Value = hippocampusFullStatus;
                    cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = hippocampus_id;
                    cmd.ExecuteNonQuery();

                    sb = new StringBuilder();
                    sb.Append("SELECT SUM(score.health_score + score.full_score + score.water_score) as fish_score ");
                    sb.Append("FROM fishbowl_pets_score score, fishbowl_pets fish ");
                    sb.Append("WHERE score.fish_id = fish.id AND fish_id = @FISH_ID");

                    string get_fish_score = sb.ToString();

                    //string get_fish_score = "SELECT TOP 1 (score.health_score + score.full_score) as fish_score FROM fishbowl_pets_score score, fishbowl_pets fish WHERE score.fish_id = fish.id AND fish_id = @FISH_ID ORDER BY score.create_time DESC";

                    string update_star = "UPDATE fishbowl_pets SET stars = @STARS WHERE id = @FISH_ID";
                    object obj_score;
                    // =======================
                    // 取出分數 anemome
                    // =======================
                    if (anemone_id > 0)
                    {
                        cmd.CommandText = get_fish_score;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = anemone_id;
                        obj_score = cmd.ExecuteScalar();
                        if (!(obj_score is DBNull))
                        {
                            anemoneScore = System.Convert.ToInt32(obj_score);
                        }
                    }

                    // =======================
                    // 取出分數 bream
                    // =======================
                    if (bream_id > 0)
                    {
                        cmd.CommandText = get_fish_score;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = bream_id;
                        obj_score = cmd.ExecuteScalar();
                        if (!(obj_score is DBNull))
                        {
                            breamScore = System.Convert.ToInt32(obj_score);
                        }
                    }

                    // =======================
                    // 取出分數 clownfish
                    // =======================
                    if (clownfish_id > 0)
                    {
                        cmd.CommandText = get_fish_score;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = clownfish_id;

                        obj_score = cmd.ExecuteScalar();
                        if (!(obj_score is DBNull))
                        {
                            clownfishScore = System.Convert.ToInt32(obj_score);
                        }
                    }

                    // =======================
                    // 取出分數 hippocampus
                    // =======================
                    if (hippocampus_id > 0)
                    {
                        cmd.CommandText = get_fish_score;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = hippocampus_id;
                        obj_score = cmd.ExecuteScalar();
                        if (!(obj_score is DBNull))
                        {
                            hippocampusScore = System.Convert.ToInt32(obj_score);
                        }
                    }

                    //string get_fish_score_history = "SELECT SUM(score.health_score + score.full_score) FROM fishbowl_pets_score score, fishbowl_pets fish WHERE score.fish_id = fish.id AND fish_id = @FISH_ID";
                    //string get_fish_score_history = "SELECT SUM(score.health_score + score.full_score) FROM fishbowl_pets_score score, fishbowl_pets fish WHERE score.fish_id = fish.id AND fish_id = @FISH_ID";

                    sb = new StringBuilder();
                    sb.Append("SELECT SUM(score.health_score + score.full_score + score.water_score) ");
                    sb.Append("FROM fishbowl_pets_score score, fishbowl_pets fish ");
                    sb.Append("WHERE score.fish_id = fish.id AND fish.fish_type = @FISH_TYPE AND fish.account_id=@ACCOUNT_ID");

                    //string get_fish_score_history = "SELECT SUM(score.health_score + score.full_score) FROM fishbowl_pets_score score, fishbowl_pets fish WHERE score.fish_id = fish.id AND fish.fish_type = @FISH_TYPE AND fish.account_id=@ACCOUNT_ID";
                    string get_fish_score_history = sb.ToString();
                    //Response.Write(get_fish_score_history);

                    //string get_fish_score_history = "SELECT SUM(score.health_score + score.full_score) FROM fishbowl_pets_score score, fishbowl_pets fish WHERE score.fish_id = fish.id AND fish.id=@FISH_ID AND fish.account_id = @ACCOUNT_ID AND fish.fish_type = (SELECT fish_type FROM fishbowl_pets WHERE id=@FISH_ID)";

                    object obj_score_history;
                    // =======================
                    // 取出歷史分數 anemome
                    // =======================
                    cmd.CommandText = get_fish_score_history;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 1;
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    obj_score_history = cmd.ExecuteScalar();

                    if (!(obj_score_history is DBNull))
                    {
                        if (anemoneScore != anemoneScoreHistory)
                        {
                            anemoneScoreHistory = System.Convert.ToInt32(obj_score_history) - anemoneScore;
                        }
                        else
                        {
                            anemoneScoreHistory = System.Convert.ToInt32(obj_score_history);
                        }
                    }

                    // 如果已經是無敵狀態...
                    if (anemoneStatus == 2)
                    {
                        anemoneStar = FishBowlUtil.calcStars(1, System.Convert.ToInt32(anemoneScore) + System.Convert.ToInt32(anemoneScoreHistory));

                        // 更新回去
                        cmd.CommandText = update_star;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = anemone_id;
                        cmd.Parameters.Add("@STARS", SqlDbType.Int, 4).Value = anemoneStar;
                        cmd.ExecuteNonQuery();
                    }

                    // =======================
                    // 取出歷史分數 bream
                    // =======================
                    cmd.CommandText = get_fish_score_history;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 2;
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    obj_score_history = cmd.ExecuteScalar();
                    if (!(obj_score_history is DBNull))
                    {
                        if (breamScore != breamScoreHistory)
                        {
                            breamScoreHistory = System.Convert.ToInt32(obj_score_history) - breamScore;
                        }
                        else
                        {
                            breamScoreHistory = System.Convert.ToInt32(obj_score_history);
                        }
                    }

                    // 如果已經是無敵狀態...
                    if (breamStatus == 2)
                    {
                        breamStar = FishBowlUtil.calcStars(2, System.Convert.ToInt32(breamScore) + System.Convert.ToInt32(breamScoreHistory));

                        // 更新回去
                        cmd.CommandText = update_star;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = bream_id;
                        cmd.Parameters.Add("@STARS", SqlDbType.Int, 4).Value = breamStar;
                        cmd.ExecuteNonQuery();
                    }

                    // =======================
                    // 取出歷史分數 clownfish
                    // =======================
                    cmd.CommandText = get_fish_score_history;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 3;
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    obj_score_history = cmd.ExecuteScalar();
                    if (!(obj_score_history is DBNull))
                    {
                        if (clownfishScore != clownfishScoreHistory)
                        {
                            clownfishScoreHistory = System.Convert.ToInt32(obj_score_history) - clownfishScore;
                        }
                        else
                        {
                            clownfishScoreHistory = System.Convert.ToInt32(obj_score_history);
                        }
                    }

                    // 如果已經是無敵狀態...
                    if (clownfishStatus == 2)
                    {
                        clownfishStar = FishBowlUtil.calcStars(3, System.Convert.ToInt32(clownfishScore) + System.Convert.ToInt32(clownfishScoreHistory));

                        // 更新回去
                        cmd.CommandText = update_star;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = clownfish_id;
                        cmd.Parameters.Add("@STARS", SqlDbType.Int, 4).Value = clownfishStar;
                        cmd.ExecuteNonQuery();
                    }

                    // =======================
                    // 取出歷史分數 hippocampus
                    // =======================
                    cmd.CommandText = get_fish_score_history;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@FISH_TYPE", SqlDbType.Int, 4).Value = 4;
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    obj_score_history = cmd.ExecuteScalar();
                    if (!(obj_score_history is DBNull))
                    {
                        if (hippocampusScore != hippocampusScoreHistory)
                        {
                            hippocampusScoreHistory = System.Convert.ToInt32(obj_score_history) - hippocampusScore;
                        }
                        else
                        {
                            hippocampusScoreHistory = System.Convert.ToInt32(obj_score_history);
                        }
                    }

                    // 如果已經是無敵狀態...
                    if (hippocampusStatus == 2)
                    {
                        hippocampusStar = FishBowlUtil.calcStars(4, System.Convert.ToInt32(hippocampusScore) + System.Convert.ToInt32(hippocampusScoreHistory));

                        // 更新回去
                        cmd.CommandText = update_star;
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add("@FISH_ID", SqlDbType.Int, 4).Value = hippocampus_id;
                        cmd.Parameters.Add("@STARS", SqlDbType.Int, 4).Value = hippocampusStar;
                        cmd.ExecuteNonQuery();
                    }

                    // 加總
                    intScore += System.Convert.ToInt32(anemoneScore + anemoneScoreHistory + breamScore + breamScoreHistory + clownfishScore + clownfishScoreHistory + hippocampusScore + hippocampusScoreHistory);

                    // 取得總星數
                    string sql = "SELECT SUM(stars) s FROM fishbowl_pets WHERE account_id = @ACCOUNT_ID";
                    cmd.CommandText = sql;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;

                    cmd.ExecuteScalar();
                    object objStars = cmd.ExecuteScalar();
                    if (!(objStars is DBNull))
                    {
                        intStars = System.Convert.ToInt32(objStars);
                    }

                    // update飽食度跟健康度
                    string sql_update_fish_status = "UPDATE fishbowl_pets SET health_status = (health_status + health_status_minus), health_status_minus = 0, full_status = (full_status + full_status_minus), full_status_minus = 0 WHERE account_id = @ACCOUNT_ID AND is_active = 1 AND status = 1";
                    cmd.CommandText = sql_update_fish_status;
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;
                    cmd.ExecuteNonQuery();

                    // =====填資料=====
                    login_id.Text = _login_id;
                    game_key.Text = _game_key;
                    money.Text = intMoney.ToString();
                    stars.Text = intStars.ToString();
                    total_score.Text = intScore.ToString();
                    water_status.Text = intWaterstatus.ToString();
                    play_time.Text = intPlayTime.ToString();

                    // Anemone
                    anemone_birth_day.Text = anemoneBirthDay.ToString();
                    anemone_full_status.Text = anemoneFullStatus.ToString();
                    anemone_health_status.Text = anemoneHealthStatus.ToString();
                    anemone_score.Text = anemoneScore.ToString();
                    anemone_history_score.Text = anemoneScoreHistory.ToString();
                    anemone_id_html.Text = anemone_id.ToString();
                    anemone_stars.Text = anemoneStar.ToString();
                    anemone_status.Text = anemoneStatus.ToString();

                    // Bream
                    bream_birth_day.Text = breamBirthDay.ToString();
                    bream_full_status.Text = breamFullStatus.ToString();
                    bream_health_status.Text = breamHealthStatus.ToString();
                    bream_score.Text = breamScore.ToString();
                    bream_history_score.Text = breamScoreHistory.ToString();
                    bream_id_html.Text = bream_id.ToString();
                    bream_stars.Text = breamStar.ToString();
                    bream_status.Text = breamStatus.ToString();

                    // Clownfish
                    clownfish_birth_day.Text = clownfishBirthDay.ToString();
                    clownfish_full_status.Text = clownfishFullStatus.ToString();
                    clownfish_health_status.Text = clownfishHealthStatus.ToString();
                    clownfish_score.Text = clownfishScore.ToString();
                    clownfish_history_score.Text = clownfishScoreHistory.ToString();
                    clownfish_id_html.Text = clownfish_id.ToString();
                    clownfish_stars.Text = clownfishStar.ToString();
                    clownfish_status.Text = clownfishStatus.ToString();

                    // Hippocampus
                    hippocampus_birth_day.Text = hippocampusBirthDay.ToString();
                    hippocampus_full_status.Text = hippocampusFullStatus.ToString();
                    hippocampus_health_status.Text = hippocampusHealthStatus.ToString();
                    hippocampus_score.Text = hippocampusScore.ToString();
                    hippocampus_history_score.Text = hippocampusScoreHistory.ToString();
                    hippocampus_id_html.Text = hippocampus_id.ToString();
                    hippocampus_stars.Text = hippocampusStar.ToString();
                    hippocampus_status.Text = hippocampusStatus.ToString();

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
