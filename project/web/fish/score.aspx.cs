using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using log4net;

public partial class score : System.Web.UI.Page
{
    private string login_id = "";
    private string game_key = "";
    private string key = "";
    private StringBuilder sb;
    private SqlCommand cmd;
    private SqlDataReader dr;
    ILog log4 = LogManager.GetLogger("FishBowl");

    protected void Page_Load(object sender, EventArgs e)
    {
        log4net.Config.XmlConfigurator.Configure();
        login_id = Request.Form["login_id"].ToString();
        game_key = Request.Form["game_key"].ToString();
        key = Request.QueryString["key"].ToString();

        if (!FishBowlUtil.isMatchedKey(login_id, game_key, key))
        {
            Response.Write("<script>alert('請重新登入!');location.href='/';</script>");
            Response.End();
            return;
        }
        else
        {
            // 取得全部排行
            all_score.Text = getAllScoreRank();

            // 取得個人排行
            personal_score.Text = getPersonalRank();
        }
    }

    /**
     * 個人排行
     */
    private string getPersonalRank()
    {
        string list = "";

        sb = new StringBuilder();
        sb.Append(" SELECT a.account_id,case when e.stars is NULL then 0 else e.stars end as TOTAL_STAR,b.realname as REALNAME,b.nickname as NICKNAME, b.login_id, ");
        sb.Append(" case when d.fish_score is NULL then 0 ");
        sb.Append(" else fish_score ");
        sb.Append(" end as TOTAL_SCORE, ");
        sb.Append(" case when f.SUCCESS_FISH_COUNT is NULL then 0 ");
        sb.Append(" else f.SUCCESS_FISH_COUNT ");
        sb.Append(" end as SUCCESS_FISH_COUNT ");
        sb.Append(" FROM fishbowl_environment a  ");
        sb.Append(" LEFT JOIN fishbowl_account b ");
        sb.Append(" on a.account_id=b.id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT SUM(temp_score) as fish_score,account_id FROM ");
        sb.Append(" (  ");
        sb.Append(" 	SELECT (z.health_score + z.full_score + z.water_score) AS temp_score,y.account_id  ");
        sb.Append(" 	FROM fishbowl_pets_score z ");
        sb.Append(" 	LEFT JOIN fishbowl_pets y ");
        sb.Append(" 	on z.fish_id=y.id ");
        sb.Append(" ) as u  ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) d ");
        sb.Append(" on a.account_id=d.account_id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT account_id,SUM(stars) as stars ");
        sb.Append(" FROM fishbowl_pets ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) e ");
        sb.Append(" on a.account_id=e.account_id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT count(distinct fish_type) as SUCCESS_FISH_COUNT,account_id  ");
        sb.Append(" FROM fishbowl_pets ");
        sb.Append(" WHERE status=2 and is_active=1 ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) f ");
        sb.Append(" on a.account_id=f.account_id ");
        sb.Append(" WHERE b.login_id=@LOGIN_ID ORDER BY SUCCESS_FISH_COUNT DESC,TOTAL_STAR DESC, TOTAL_SCORE DESC, b.create_time");

        SqlConnection conn = SQLDB.createConnection();
        try
        {
            using (cmd = new SqlCommand(sb.ToString(), conn))
            {
                cmd.CommandText = sb.ToString();
                cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = login_id;

                dr = cmd.ExecuteReader();

                StringBuilder sb_personal = new StringBuilder();
                string success_fishes = "";
                string stars = "";
                string real_name = "";
                string nick_name = "";

                if (dr.Read())
                {
                    if (dr["TOTAL_STAR"] is DBNull)
                    {
                        stars = "0";
                    }
                    else
                    {
                        stars = System.Convert.ToString(dr["TOTAL_STAR"]);
                    }
                    real_name = System.Convert.ToString(dr["REALNAME"]);
                    nick_name = System.Convert.ToString(dr["NICKNAME"]);
                    success_fishes = System.Convert.ToString(dr["SUCCESS_FISH_COUNT"]);

                    if (nick_name != "")
                    {
                        sb_personal.Append("<username>" + nick_name + "</username>");
                    }
                    else
                    {
                        sb_personal.Append("<username>" + replaceRealName(real_name) + "</username>");
                    }

                    sb_personal.Append("<score>" + System.Convert.ToString(dr["TOTAL_SCORE"]) + "</score>");
                    sb_personal.Append("<stars>" + stars + "</stars>");
                    sb_personal.Append("<rank>" + getRanking(System.Convert.ToInt32(dr["account_id"])) + "</rank>");
                    sb_personal.Append("<success_fishes>" + success_fishes + "</success_fishes>");
                    sb_personal.Append("<has_lottery>" + checkLotteryChance(System.Convert.ToInt32(success_fishes), System.Convert.ToInt32(stars)) + "</has_lottery>");
                }

                list = sb_personal.ToString();
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

        return list;
    }

    /**
     * 全部排行
     */
    private string getAllScoreRank()
    {
        sb = new StringBuilder();
        sb.Append(" SELECT TOP 10 a.account_id,case when e.stars is NULL then 0 else e.stars end as TOTAL_STAR,b.realname as REALNAME,b.nickname as NICKNAME, b.login_id, ");
        sb.Append(" case when d.fish_score is NULL then 0 ");
        sb.Append(" else fish_score ");
        sb.Append(" end as TOTAL_SCORE, ");
        sb.Append(" case when f.SUCCESS_FISH_COUNT is NULL then 0 ");
        sb.Append(" else f.SUCCESS_FISH_COUNT ");
        sb.Append(" end as SUCCESS_FISH_COUNT ");
        sb.Append(" FROM fishbowl_environment a  ");
        sb.Append(" LEFT JOIN fishbowl_account b ");
        sb.Append(" on a.account_id=b.id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT SUM(temp_score) as fish_score,account_id FROM ");
        sb.Append(" (  ");
        sb.Append(" 	SELECT (z.health_score + z.full_score + z.water_score) AS temp_score,y.account_id  ");
        sb.Append(" 	FROM fishbowl_pets_score z ");
        sb.Append(" 	LEFT JOIN fishbowl_pets y ");
        sb.Append(" 	on z.fish_id=y.id ");
        sb.Append(" ) as u  ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) d ");
        sb.Append(" on a.account_id=d.account_id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT account_id,SUM(stars) as stars ");
        sb.Append(" FROM fishbowl_pets ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) e ");
        sb.Append(" on a.account_id=e.account_id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT count(distinct fish_type) as SUCCESS_FISH_COUNT,account_id  ");
        sb.Append(" FROM fishbowl_pets ");
        sb.Append(" WHERE status=2 and is_active=1 ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) f ");
        sb.Append(" on a.account_id=f.account_id ");
        sb.Append(" order by SUCCESS_FISH_COUNT DESC,TOTAL_STAR DESC, TOTAL_SCORE DESC, b.create_time");

        SqlConnection conn = SQLDB.createConnection();
        try
        {
            using (cmd = new SqlCommand(sb.ToString(), conn))
            {
                using (dr = cmd.ExecuteReader())
                {
                    sb = new StringBuilder();

                    string real_name = "";
                    string nick_name = "";
                    string account_id = "";
                    string success_fishes = "";
                    string stars = "";
                    if (dr.HasRows)
                    {
                        while (dr.Read())
                        {
                            real_name = System.Convert.ToString(dr["REALNAME"]);
                            nick_name = System.Convert.ToString(dr["NICKNAME"]);
                            account_id = System.Convert.ToString(dr["account_id"]);
                            success_fishes = System.Convert.ToString(dr["SUCCESS_FISH_COUNT"]);
                            if (dr["TOTAL_STAR"] is DBNull)
                            {
                                stars = "0";
                            }
                            else
                            {
                                stars = System.Convert.ToString(dr["TOTAL_STAR"]);
                            }

                            sb.Append("<list>");
                            if (nick_name != "")
                            {
                                sb.Append("<username>" + nick_name + "</username>");
                            }
                            else
                            {
                                sb.Append("<username>" + replaceRealName(real_name) + "</username>");
                            }
                            sb.Append("<login_id>" + System.Convert.ToString(dr["login_id"]) + "</login_id>");
                            sb.Append("<success_fishes>" + success_fishes + "</success_fishes>");
                            sb.Append("<stars>" + stars + "</stars>");
                            sb.Append("<score>" + System.Convert.ToString(dr["TOTAL_SCORE"]) + "</score>");
                            sb.Append("</list>");
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
        return sb.ToString();
    }

    /**
     * 把真實姓名馬賽克
     */
    private string replaceRealName(string oldname)
    {
        return (oldname.Substring(0, 1) + "＊" + oldname.Substring(oldname.Trim().Length - 1, 1));
    }

    /**
     * 是否有抽獎資格
     */
    private string checkLotteryChance(int fishes, int stars)
    {
        // 曾經成功養過1種動物，且累計星等達34顆星以上
        // 2009.10.15 改成1種動物，10顆星以上
        if (fishes > 0 && stars >= 10)
        {
            return "1";
        }
        else
        {
            return "0";
        }
    }

    /**
     * 取得個人排名
     */
    private string getRanking(int account_id)
    {
        int ranking = 0;

        SqlConnection conn = SQLDB.createConnection();
        
        sb = new StringBuilder();
        sb.Append(" DECLARE @RankingTable TABLE (ranking int, account_id int, stars int, scores int, success_fishes int) ");
        sb.Append(" INSERT INTO @RankingTable (ranking, account_id, stars, scores, success_fishes) ");
        sb.Append(" SELECT row_number() OVER(ORDER BY f.SUCCESS_FISH_COUNT DESC,e.stars DESC, d.fish_score DESC, b.create_time) as ranking, a.account_id,e.stars as TOTAL_STAR, ");
        sb.Append(" case when d.fish_score is NULL then 0 ");
        sb.Append(" else fish_score ");
        sb.Append(" end as TOTAL_SCORE, ");
        sb.Append(" case when f.SUCCESS_FISH_COUNT is NULL then 0 ");
        sb.Append(" else f.SUCCESS_FISH_COUNT ");
        sb.Append(" end as SUCCESS_FISH_COUNT ");
        sb.Append(" FROM fishbowl_environment a  ");
        sb.Append(" LEFT JOIN fishbowl_account b ");
        sb.Append(" on a.account_id=b.id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT SUM(temp_score) as fish_score,account_id FROM ");
        sb.Append(" (  ");
        sb.Append(" 	SELECT (z.health_score + z.full_score + z.water_score) AS temp_score,y.account_id  ");
        sb.Append(" 	FROM fishbowl_pets_score z ");
        sb.Append(" 	LEFT JOIN fishbowl_pets y ");
        sb.Append(" 	on z.fish_id=y.id ");
        sb.Append(" ) as u  ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) d ");
        sb.Append(" on a.account_id=d.account_id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT account_id,SUM(stars) as stars ");
        sb.Append(" FROM fishbowl_pets ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) e ");
        sb.Append(" on a.account_id=e.account_id ");
        sb.Append(" LEFT JOIN ( ");
        sb.Append(" SELECT count(distinct fish_type) as SUCCESS_FISH_COUNT,account_id  ");
        sb.Append(" FROM fishbowl_pets ");
        sb.Append(" WHERE status=2 and is_active=1 ");
        sb.Append(" GROUP BY account_id ");
        sb.Append(" ) f ");
        sb.Append(" on a.account_id=f.account_id ");
        sb.Append(" SELECT ranking FROM @RankingTable WHERE account_id=@ACCOUNT_ID");

        try
        {
            using (SqlCommand cmd = new SqlCommand(sb.ToString(), conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id;

                object o = cmd.ExecuteScalar();
                if (!(o is DBNull))
                {
                    ranking = Convert.ToInt32(o);
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
        return ranking.ToString();
    }
}
