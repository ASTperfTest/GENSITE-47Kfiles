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
using GSS.Vitals.COA.Data;

public class TreasureHunt
{
    private SqlCommand cmd;
    private SqlDataReader dr;
    public string loginId = "";
    Activity activity;
    private int accountId;
    private ILog log4 = LogManager.GetLogger("DR");

    public TreasureHunt(string login_id)
    {
        log4net.Config.XmlConfigurator.Configure();

        if (!string.IsNullOrEmpty(login_id) )
        {
            loginId = login_id.Trim();
            accountId = account_id;
        }
        activity = getActivity();
    }

    public void SetActivity(int actId)
    {
        activity = new Activity();
        activity.Id = actId;
    }

    class Activity
    {
        int m_Id;
        string m_LIMIT_PAGES;
        int m_Activity_Type;
        public int Id
        {
            get { return m_Id; }
            set { m_Id = value; }
        }
        public string LIMIT_PAGES
        {
            get { return m_LIMIT_PAGES; }
            set { m_LIMIT_PAGES = value; }
        }
        public int Activity_Type
        {
            get { return m_Activity_Type; }
            set { m_Activity_Type = value; }
        }

    }



    /**
     * 登入時更新account
     */
    public void loginUpdate(string nickname, string realname, string email)
    {
        string sql = "SELECT ACCOUNT_ID,REALNAME,NICKNAME,EMAIL FROM ACCOUNT WHERE LOGIN_ID = @LOGIN_ID";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd_update = new SqlCommand(sql, conn))
            {
                cmd_update.Parameters.Clear();
                cmd_update.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = loginId;
                int account_id = 0;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd_update.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            // 有資料，檢查資料是否變更並更新
                            account_id = Convert.ToInt32(dr["ACCOUNT_ID"]);
                            string oldNickname = Convert.ToString(dr["NICKNAME"]);
                            string oldRealname = Convert.ToString(dr["REALNAME"]);
                            string oldEmail = Convert.ToString(dr["EMAIL"]);
                            dr.Close();
                            if (nickname.CompareTo(oldNickname) != 0 || realname.CompareTo(oldRealname) != 0 || email.CompareTo(oldEmail) != 0)
                            {
                                using (SqlCommand cmd_login = new SqlCommand())
                                {
                                    sql = "UPDATE ACCOUNT SET REALNAME=@REAL_NAME, NICKNAME=@NICK_NAME,EMAIL=@EMAIL, MODIFY_DATE=@MODIFY_DATE WHERE ACCOUNT_ID=@ACCOUNT_ID";

                                    cmd_login.CommandText = sql;
                                    cmd_login.Connection = conn;
                                    cmd_login.Parameters.Clear();
                                    cmd_login.Parameters.Add("@REAL_NAME", SqlDbType.NVarChar, 200).Value = realname;
                                    cmd_login.Parameters.Add("@NICK_NAME", SqlDbType.NVarChar, 200).Value = nickname;
                                    cmd_login.Parameters.Add("@EMAIL", SqlDbType.NVarChar, 200).Value = email;
                                    cmd_login.Parameters.Add("@MODIFY_DATE", SqlDbType.DateTime).Value = DateTime.Now;
                                    cmd_login.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = account_id;

                                    cmd_login.ExecuteNonQuery();
                                }
                            }


                            //log(account_id, "Login");
                        }
                        else
                        {
                            // 無資料，寫入
                            dr.Close();
                            sql = "INSERT INTO ACCOUNT(LOGIN_ID, CREATE_DATE, REALNAME, NICKNAME, EMAIL) ";
                            sql += " VALUES(@ID, GETDATE() , @REAL_NAME, @NICK_NAME, @EMAIL) ";
                            using (SqlCommand cmd_insert = new SqlCommand(sql, conn))
                            {
                                cmd_insert.Parameters.Clear();
                                cmd_insert.Parameters.Add("@ID", SqlDbType.VarChar, 50).Value = loginId;
                                cmd_insert.Parameters.Add("@REAL_NAME", SqlDbType.VarChar, 200).Value = realname;
                                cmd_insert.Parameters.Add("@NICK_NAME", SqlDbType.VarChar, 200).Value = nickname;
                                cmd_insert.Parameters.Add("@EMAIL", SqlDbType.VarChar, 200).Value = email;
                                cmd_insert.ExecuteNonQuery();

                                cmd_insert.CommandText = "Select @@Identity";
                                object o = cmd_insert.ExecuteScalar();
                                if (!(o is DBNull))
                                {
                                    account_id = Convert.ToInt32(o);
                                }
                            }
                        }
                    }
                    if (account_id != 0)
                        CreateProbabilityAndPackage(account_id);
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
     * 新增未開到寶物次數
     */
    private void CreateProbabilityAndPackage(int newId)
    {
        string sql = "";

        sql = @"
            BEGIN TRAN
            if not exists(SELECT * FROM PROBABILITY WHERE ACCOUNT_ID = @ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID)
            BEGIN 
	            INSERT INTO PROBABILITY(ACCOUNT_ID, ACTIVITY_ID, TIMES) 
	            VALUES (@ACCOUNT_ID, @ACTIVITY_ID, @TIMES) 
            END
            IF NOT EXISTS ( SELECT * FROM  USER_SEARCH_TIMER WHERE ACCOUNT_ID = @ACCOUNT_ID)
            BEGIN 
	             INSERT INTO  USER_SEARCH_TIMER (ACCOUNT_ID,LAST_SEARCH_TIME) 
	             VALUES (@ACCOUNT_ID, DateAdd(d,-1,GETDATE()))
            END
            IF NOT EXISTS (SELECT * FROM USERS_PACKAGE  WHERE ACCOUNT_ID = @ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID)
            BEGIN
	            INSERT INTO USERS_PACKAGE(ACCOUNT_ID,ACTIVITY_ID,TREASURE_ID,PIECE)
	            SELECT @ACCOUNT_ID,@ACTIVITY_ID,TREASURE_ID,0 FROM TREASURE
	            WHERE ACTIVITY_ID = @ACTIVITY_ID 
            END
            commit
        ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = newId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                cmd.Parameters.Add("@TIMES", SqlDbType.Int).Value = 1;
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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

    private Activity getActivity()
    {

        Activity act = new Activity();
        string sql = "SELECT ACTIVITY_ID,LIMIT_PAGES,ACTIVITY_TYPE FROM ACTIVITY WHERE GETDATE() BETWEEN ONLINE_DATE AND OFFLINE_DATE";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            act.Id = Convert.ToInt32(dr["ACTIVITY_ID"]);
                            act.LIMIT_PAGES = Convert.ToString(dr["LIMIT_PAGES"]);
                            act.Activity_Type =  Convert.ToInt32(dr["ACTIVITY_TYPE"]);
                        }
                        else
                        {
                            act = null;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return act;

    }
    /**
     * 取得account_id
     */
    public int account_id
    {
        get
        {
            int id = 0;

            string sql = "SELECT ACCOUNT_ID FROM ACCOUNT WHERE login_id = @LOGIN_ID";

            using (SqlConnection conn = TreasureHuntDB.createConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@LOGIN_ID", SqlDbType.VarChar, 50).Value = loginId;
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
                        throw ex;
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
    public int getactivity_id()
    {
        return activity.Id;
    }
    public int getintervall_days()
    {
        return intervall_days;
    }

    public int intervall_days
    {
        get
        {
            int intervallDays = 0;
            string temp = GetConfigData("INTERVAL_DAYS");
            if (!string.IsNullOrEmpty(temp))
                intervallDays = Convert.ToInt32(temp);
            return intervallDays;
        }
    }

    public string reciprocal_time
    {
        get
        {
            string returnString = GetConfigData("RECIPROCAL_TIME");
            return returnString;
        }
    }

    public bool checkActivity(string url,string param,string referer_url)
    {
       // if (activity == null || !activity.LIMIT_PAGES.Contains(url) && (!string.IsNullOrEmpty(url)))
        if (activity != null && (!string.IsNullOrEmpty(url)))
        {
            if (CheckActivityBytype(url, param, referer_url))
            return true;
        }
        return false;
    }

    private bool CheckActivityBytype(string url, string param, string referer_url)
    {
        switch (activity.Activity_Type)
        {
            case 1:
                return (ContainsJijsawUrl(url + param, referer_url));
            case 2:
                return (IsActivityPage(url) || ContainsJijsawUrl(url + param, referer_url));
            default:
                return IsActivityPage(url);
        }
    }

    private bool IsActivityPage(string url)
    {
        string[] temp = activity.LIMIT_PAGES.Split(';');
        foreach (string s in temp)
        {
            if (url.Contains(s))
                return true;
        }
        return false;
    }

    private bool ContainsJijsawUrl(string url, string referer_url)
    {
        string[] activityPage = activity.LIMIT_PAGES.Split(';');
        bool isActivityPage= false;
        if (!string.IsNullOrEmpty(referer_url))
        foreach (string s in activityPage)
        {
            if (referer_url.Contains(s))
                isActivityPage = true;
        }
        if (isActivityPage == false)
            return false;
        
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        string documenttypeandname = ""; 
        string sql = "";
        if (url.Contains("/jigsaw2010/detail.aspx") || url.Contains("%2fjigsaw2010%2fdetail.aspx"))
        {
            int articleId = getDocumentId(url);
            documenttypeandname = articleId.ToString();
            sql = @"
                IF EXISTS (
		                    SELECT * FROM 
	                             CuDTGeneric
                            where FCTUPUBLIC = 'Y' AND ICUITEM = @XURL
                )
	                SELECT 1
            ";
        }
        else
        {
            documenttypeandname = "%" + GetDocumentIdAndType(url)+"%";
            sql = @"     
                IF EXISTS (
		                    SELECT * FROM 
	                            KnowledgeJigsaw 
                            where status = 'Y' AND PATH like @XURL
                )
	                SELECT 1
        ";
        }
        var dt = SqlHelper.GetDataTable("GSSConnString", sql,
             DbProviderFactories.CreateParameter("GSSConnString", "@XURL", "@XURL", documenttypeandname));
        if (dt.Rows.Count > 0)
            return true;
        return false;
    }

    public string getActivityPages()
    {
        return activity.LIMIT_PAGES;
    }

    public string getCheatMode
    {
        get
        {
            string returnString = GetConfigData("CHEAT_MODE");
            return returnString;
        }
        set
        {
            InsertOrUpdateConfigData("CHEAT_MODE", value);
        }
    }

    public string getCheatModeEnd
    {
        get
        {
            string returnString = GetConfigData("CHEAT_MODE_END");
            return returnString;
        }
        set
        {
            InsertOrUpdateConfigData("CHEAT_MODE_END", value);
        }
    }

    public string getLotteryStartDate
    {
        get
        {
            string returnString = GetConfigData("LOTTERY_START");
            return returnString;
        }
        set
        {
            InsertOrUpdateConfigData("LOTTERY_START", value);
        }
    }
    public string getLotteryEndDate
    {
        get
        {
            string returnString = GetConfigData("LOTTERY_END");
            return returnString;
        }
        set
        {
            InsertOrUpdateConfigData("LOTTERY_END", value);
        }
    }

    public Dictionary<string, string> GetUserData(string userLoginid)
    {
        loginId = userLoginid.Trim();
        return GetUserData(account_id);
    }

    public Dictionary<string, string> GetUserData(int id)
    {
        Dictionary<string, string> returnValue = new Dictionary<string, string>();
        string sql = "SELECT * FROM ACCOUNT WHERE ACCOUNT_ID = @ACCOUNT_ID ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = id;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            returnValue.Add("loginId", dr["LOGIN_ID"].ToString());
                            returnValue.Add("name", replaceRealName(dr["NICKNAME"].ToString(), dr["REALNAME"].ToString()));
                            returnValue.Add("mail", dr["EMAIL"].ToString());
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
        return returnValue;
    }

    public string GetConfigData(string configId)
    {
        string configValue = "";
        string sql = "SELECT VALUE FROM CONFIG_DATA WHERE CONFIG_ID = @CONFIG_ID AND ACTIVITY_ID = @ACTIVITY_ID ";

        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                cmd.Parameters.Add("@CONFIG_ID", SqlDbType.NVarChar, 50).Value = configId;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        configValue = Convert.ToString(o);
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return configValue;

    }

    public bool InsertOrUpdateConfigData(string configId , string newValue)
    {
        string sql = "IF EXISTS(SELECT VALUE FROM dbo.CONFIG_DATA WHERE CONFIG_ID=@CONFIG_ID AND ACTIVITY_ID=@ACTIVITY_ID ) ";
        sql += " BEGIN ";
        sql += " UPDATE CONFIG_DATA SET [VALUE]=@NEWVALUE WHERE CONFIG_ID=@CONFIG_ID AND ACTIVITY_ID=@ACTIVITY_ID ";
        sql += " END ELSE BEGIN ";
	    sql += " INSERT INTO dbo.CONFIG_DATA( CONFIG_ID, ACTIVITY_ID, [VALUE] ) ";
        sql += " VALUES  ( @CONFIG_ID,@ACTIVITY_ID,@NEWVALUE) ";
        sql += " END ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                cmd.Parameters.Add("@CONFIG_ID", SqlDbType.NVarChar,50).Value = configId;
                cmd.Parameters.Add("@NEWVALUE", SqlDbType.NVarChar, 50).Value = newValue;
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    ////throw ex;
                    log4.Error( ex.ToString());
                    return false;
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
        return true;
    }

    /**
    * 檢查閱讀紀錄
    * 閱讀的文章要3天內未閱覽才可算是能進行遊戲的文章
    */
    public int checkReadDocument(string url)
    {
        int returnValue = 0;
        string sql = " {0}  ";
        string temp = "";
        int q = intervall_days;
        string documentTypeAndName = GetDocumentIdAndType(url);
        for (int i = 0; i < q; i++)
        {
            temp = " IF EXISTS( ";
            temp += " SELECT * FROM READ_DOCUMENT_RECORD ";
            temp += " WHERE ACTIVITY_ID = @ACTIVITY_ID AND ACCOUNT_ID = @ACCOUNT_ID AND URL LIKE @URL AND ";
            temp += " CONVERT(varchar(12) , DateAdd(\"d\",@INTERVAL_DAYS" + (q - i).ToString() + ",MODIFY_DATE) , 111 ) >  CONVERT(varchar(12) , GETDATE() , 111 ) )";
            temp += " {0} ";
            temp += " ELSE SELECT " + i.ToString() + "  ";
            sql = string.Format(sql, temp);
        }
        sql = string.Format(sql, "SELECT " + q.ToString());
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                for (int i = 0; i < q; i++)
                {
                    cmd.Parameters.Add("@INTERVAL_DAYS" + (q - i).ToString(), SqlDbType.Int).Value = q - i;
                }
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = account_id;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                cmd.Parameters.Add("@URL", SqlDbType.NVarChar, 255).Value = "%"+documentTypeAndName + "%";
                cmd.Parameters.Add("@INTERVAL_DAYS", SqlDbType.Int).Value = intervall_days;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    log4.Error("account_id:" + account_id.ToString() + "__" + url +" \n"+ ex.ToString());
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
        return returnValue;
    }


    /// <summary>
    /// 檢查最後一次倒數時間
    /// </summary>
    /// <returns></returns>
    public bool CheckLastSearchTime()
    {

        bool returnValue = false;
        string sql = " BEGIN TRAN ";
	    sql += "  IF EXISTS (SELECT * FROM USER_SEARCH_TIMER (updlock) ";
        sql += "  WHERE ACCOUNT_ID = @ACCOUNT_ID AND DateAdd(\"s\",@RECIPROCAL_TIME,LAST_SEARCH_TIME) < GETDATE()) ";
	    sql += "  begin ";
        sql += "    UPDATE USER_SEARCH_TIMER SET LAST_SEARCH_TIME = GETDATE() WHERE ACCOUNT_ID = @ACCOUNT_ID ";
		sql += "    SELECT 1 ";
	    sql += "  END ";
	    sql += "  ELSE ";
	    sql += "  begin ";
		sql += "    SELECT 0 ";
	    sql += "  END ";
        sql += "  commit ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = account_id;
                cmd.Parameters.Add("@RECIPROCAL_TIME", SqlDbType.Int).Value = reciprocal_time;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToBoolean(Convert.ToInt32(o));
                    }
                }
                catch (Exception ex)
                {
                    log4.Error("account_id:"+account_id.ToString()+"\n" + ex.ToString());
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
        return returnValue;
    }

    /**
     * 尋寶開始
     */
    public bool SearchTreasureBox(string url)
    {
        bool find = false;
        int accountId = account_id;
        ProbabilityNumber probabilityNumber = GetProbabilityNumber(accountId);
        Random rnd = new Random();
        //假設尋寶次數已經大於等於機機率分母就直接算找到以免程式發生問題
        if (probabilityNumber.Minimum >= probabilityNumber.Maximum)
        {
            find = true;
        }
        else
            if (probabilityNumber.Minimum >= rnd.Next(1, probabilityNumber.Maximum))
            {
                find = true;
            }
        return find;
    }


    //取得計算機率的分母跟分子
    private ProbabilityNumber GetProbabilityNumber(int accountId)
    {
        try
        {
            ProbabilityNumber p = new ProbabilityNumber();
            int searchTimes = GetProbabilityCount(accountId);
            int startingProbabilityPage = starting_Probability_Page;
            if (searchTimes < startingProbabilityPage)
            {
                string[] Probability = starting_Probability.Split(';');
                p.Maximum = Convert.ToInt32(Probability[Probability.Length - 1]);
                p.Minimum = Convert.ToInt32(Probability[0]);

            }
            else
            {
                p.Minimum = searchTimes;
                p.Maximum = Probability;
            }

            return p;
        }
        catch (Exception ex)
        {
            log4.Error(ex.ToString());
            throw ex;
        }
    }


    //取得使用者未尋找到寶物的次數
    private int GetProbabilityCount(int accountId)
    {
        int returnValue = 1;
        string sql = "SELECT TIMES FROM PROBABILITY WHERE ACTIVITY_ID = @ACTIVITY_ID AND ACCOUNT_ID = @ACCOUNT_ID  ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = accountId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
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
        return returnValue;
    }

    //測試用
    public int getValueee()
    {
        ProbabilityNumber p = GetProbabilityNumber(account_id);
        return p.Minimum;
    }

    //取得最高抽獎次數
    public int limit_set
    {
        get
        {
            return Convert.ToInt32(GetConfigData("LIMIT_SET"));
        }
    }

    //取得起始機率分母
    private string starting_Probability
    {
        get
        {
            return GetConfigData("STARTING_PROBABILITY");
        }
    }

    //取得超過起始機率頁數之後的機率分母
    private int Probability
    {
        get
        {
            int returnValue = 1;
            string temp = GetConfigData("PROBABILITY");
            if (!string.IsNullOrEmpty(temp))
                returnValue = Convert.ToInt32(temp);
            return returnValue;
        }
    }

    //取得起始機率的頁數
    private int starting_Probability_Page
    {
        get
        {
            int returnValue = 1;
            string temp = GetConfigData("STARTING_PROBABILITY_PAGE");
            if (!string.IsNullOrEmpty(temp))
                returnValue = Convert.ToInt32(temp);
            return returnValue;
        }
    }

    //取得每日取寶上限
    private int daily_set
    {
        get
        {
            int returnValue = 1;
            string temp = GetConfigData("DAILY_SET");
            if (!string.IsNullOrEmpty(temp))
                returnValue = Convert.ToInt32(temp);
            return returnValue;
        }
    }

    #region 讀文章取寶相關
    /// <summary>
    /// 新增修改READ_DOCUMENT_RECORD 以及取寶log
    /// </summary>
    /// <param name="url">文章網址</param>
    /// <param name="find">是否找到寶物</param>
    public int SaveReadDocimentRecordData(string url, bool find)
    {
        int treasureLogId = -1;
        string documentTypeAndName = GetDocumentIdAndType(url);
        string sql = "SELECT * FROM READ_DOCUMENT_RECORD WHERE ACTIVITY_ID = @ACTIVITY_ID AND ACCOUNT_ID = @ACCOUNT_ID AND URL like @URL ";
        ReadDocumentRecord readDocumentRecord = new ReadDocumentRecord();
        int accountId = account_id;
        int activityId = activity.Id;
        int searchCount = 1;
        int states = 0;
        string message = "";
        //有無找到寶物狀態分別
        if (!find)
        {
            states = (int)TreasureHuntState.Faild;
            message = Enum.GetName(typeof(TreasureHuntState), (int)TreasureHuntState.Faild);
            searchCount = GetProbabilityCount(accountId) + 1;
        }
        else
        {
            states = (int)TreasureHuntState.Get_Box;
            message = Enum.GetName(typeof(TreasureHuntState), (int)TreasureHuntState.Get_Box);
        }
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = accountId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activityId;
                cmd.Parameters.Add("@URL", SqlDbType.NVarChar, 255).Value = "%"+documentTypeAndName+"%";
                try
                {
                    conn.Open();

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {

                        if (dr.Read())
                        {

                            readDocumentRecord.ID = Convert.ToInt32(dr["READ_DOCUMENT_ID"]);
                            readDocumentRecord.Times = Convert.ToInt32(dr["TIMES"]) + 1;
                            dr.Close();
                            sql = " UPDATE  READ_DOCUMENT_RECORD SET TIMES=@HIT_TIMES,MODIFY_DATE=GETDATE() WHERE READ_DOCUMENT_ID = @READ_DOCUMENT_ID And ACTIVITY_ID = @ACTIVITY_ID AND ACCOUNT_ID = @ACCOUNT_ID ";
                            sql += " UPDATE PROBABILITY SET TIMES =@PROBABILITY WHERE ACTIVITY_ID = @ACTIVITY_ID AND ACCOUNT_ID = @ACCOUNT_ID ";
                            sql += " INSERT INTO TREASURE_HUNT_LOG(ACTIVITY_ID,ACCOUNT_ID,READ_DOCUMENT_ID,STATES,MESSAGE,MODIFY_DATE,CREATE_DATE) ";
                            sql += " VALUES (@ACTIVITY_ID,@ACCOUNT_ID,@READ_DOCUMENT_ID,@STATES,@MESSAGE,GETDATE(),GETDATE()) ";
                            sql += " Select @@Identity ";
                            using (SqlCommand cmd_iou = new SqlCommand(sql, conn))
                            {
                                cmd_iou.Parameters.Clear();
                                cmd_iou.CommandText = sql;
                                cmd_iou.Parameters.Add("@READ_DOCUMENT_ID", SqlDbType.Int).Value = readDocumentRecord.ID;
                                cmd_iou.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = accountId;
                                cmd_iou.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activityId;
                                cmd_iou.Parameters.Add("@HIT_TIMES", SqlDbType.Int).Value = readDocumentRecord.Times;
                                cmd_iou.Parameters.Add("@PROBABILITY", SqlDbType.Int).Value = searchCount;

                                cmd_iou.Parameters.Add("@STATES", SqlDbType.Int).Value = states;
                                cmd_iou.Parameters.Add("@MESSAGE", SqlDbType.NVarChar, 255).Value = message;
                                object o = cmd_iou.ExecuteScalar();
                                if (!(o is DBNull))
                                {
                                    treasureLogId = Convert.ToInt32(o);
                                }
                            }
                        }
                        else
                        {
                            dr.Close();
                            
                            readDocumentRecord.Url = url;
                            readDocumentRecord.CreateDate = DateTime.Now;
                            readDocumentRecord.Times = 1;
                            sql = " INSERT INTO  READ_DOCUMENT_RECORD(ACTIVITY_ID,ACCOUNT_ID,URL,TIMES,DOCUMENT_TITLE,UNIT_NAME,MODIFY_DATE,CREATE_DATE)";
                            sql += " VALUES(@ACTIVITY_ID,@ACCOUNT_ID,@URL,1,@DOCUMENT_TITLE,@UNIT_NAME,Getdate(),Getdate()) ";
                            sql += " INSERT INTO TREASURE_HUNT_LOG(ACTIVITY_ID,ACCOUNT_ID,READ_DOCUMENT_ID,STATES,MESSAGE,MODIFY_DATE,CREATE_DATE) ";
                            sql += " VALUES (@ACTIVITY_ID,@ACCOUNT_ID,@@Identity,@STATES,@MESSAGE,Getdate(),Getdate()) ";
                            sql += " UPDATE PROBABILITY SET TIMES =@PROBABILITY  WHERE ACTIVITY_ID = @ACTIVITY_ID AND ACCOUNT_ID = @ACCOUNT_ID ";
                            sql += " Select @@Identity ";
                            using (SqlCommand cmd_iou = new SqlCommand(sql, conn))
                            {
                                cmd_iou.Parameters.Clear();
                                cmd_iou.CommandText = sql;
                                cmd_iou.Parameters.Add("@READ_DOCUMENT_ID", SqlDbType.Int).Value = readDocumentRecord.ID;
                                cmd_iou.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = accountId;
                                cmd_iou.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activityId;
                                cmd_iou.Parameters.Add("@HIT_TIMES", SqlDbType.BigInt).Value = readDocumentRecord.Times;
                                cmd_iou.Parameters.Add("@PROBABILITY", SqlDbType.Int).Value = searchCount;
                                cmd_iou.Parameters.Add("@STATES", SqlDbType.Int).Value = states;
                                cmd_iou.Parameters.Add("@URL", SqlDbType.NVarChar, 255).Value = url;
                                int articleId = getDocumentId(url);
                                string documentTitle = "default";
                                if (articleId != -1)
                                    documentTitle = getDocumentTitle(articleId);
                                cmd_iou.Parameters.Add("@DOCUMENT_TITLE", SqlDbType.NVarChar, 255).Value = documentTitle;
                                cmd_iou.Parameters.Add("@UNIT_NAME", SqlDbType.NVarChar, 50).Value = GetActivityUnitName(url);
                                cmd_iou.Parameters.Add("@MESSAGE", SqlDbType.NVarChar, 255).Value = message;
                                object o = cmd_iou.ExecuteScalar();
                                if (!(o is DBNull))
                                {
                                    treasureLogId = Convert.ToInt32(o);
                                }
                            }
                        }

                    }
                    return treasureLogId;
                }
                catch (Exception ex)
                {
                    ////throw ex;
                    log4.Error("account_id:"+account_id.ToString()+"\n url=" + url + ex.ToString());
                    return treasureLogId;
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

    /// <summary>
    /// 更新取寶之後的log
    /// </summary>
    /// <param name="treasureLodId"></param>
    /// <returns></returns>
    public int UpdateTreasureHuntLog(int treasureLodId)
    {
        //活動結束前專用 直接獲取user 缺的寶物
        int treasureId = 0;
        int actId = activity.Id;
        int userId = accountId;
        if (CheatModeOpen())
        {
            treasureId = GetUserNeedsTreasure(actId, userId);
        }
        else
        {
            treasureId = GetTreasure();
        }
       
        string sql = @" 
         BEGIN TRAN
		 if exists ( select * from TREASURE_HUNT_LOG WHERE TREASURE_HUNT_LOG_ID = @TREASURE_HUNT_LOG_ID AND ACCOUNT_ID = @ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID AND STATES = @STATESGET )
		 BEGIN
			UPDATE TREASURE_HUNT_LOG SET STATES = @STATES ,MESSAGE = @MESSAGE,MODIFY_DATE = GETDATE(),TREASURE_ID=@TREASURE_ID  
			WHERE TREASURE_HUNT_LOG_ID = @TREASURE_HUNT_LOG_ID AND ACCOUNT_ID = @ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID AND STATES = @STATESGET 
			UPDATE USERS_PACKAGE SET PIECE = PIECE + 1 WHERE  ACCOUNT_ID = @ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID AND TREASURE_ID=@TREASURE_ID 
			SELECT 1;
			END
			ELSE BEGIN
				SELECT 0;
			END
         COMMIT
        ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = accountId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                cmd.Parameters.Add("@STATES", SqlDbType.Int).Value = (int)TreasureHuntState.Get_Treasure;
                cmd.Parameters.Add("@MESSAGE", SqlDbType.NVarChar, 255).Value = Enum.GetName(typeof(TreasureHuntState), (int)TreasureHuntState.Get_Treasure);
                cmd.Parameters.Add("@TREASURE_HUNT_LOG_ID", SqlDbType.Int).Value = treasureLodId;
                cmd.Parameters.Add("@TREASURE_ID", SqlDbType.Int).Value = treasureId;
                cmd.Parameters.Add("@STATESGET", SqlDbType.Int).Value = (int)TreasureHuntState.Get_Box;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        string temp = Convert.ToString(o);
						if(temp=="0")
						return 0;
							
                    }
					
                }
                catch (Exception ex)
                {
                    ////throw ex;
                    log4.Error("userId:" + userId.ToString() + "\n treasureLodId:" + treasureLodId.ToString() + "\n" + ex.ToString());
                    return 0;
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
        return treasureId;
    }

    public int RandomChangeTreasure(int userId, int actId, int treasureId, int treasureId2)
    {
        int trId =  GetCheatRandomChangeTreasure(userId,actId,treasureId,treasureId2);
        string sql = @"
                BEGIN TRAN
                DECLARE @changetreasureid int,@changetreasureid2 int,@icti int
                if not exists( SELECT * FROM  TREASURE_HUNT_LOG WHERE ACCOUNT_ID = @ACCOUNT_ID AND STATES=@STATES2 AND ACTIVITY_ID=@ACTIVITY_ID AND CONVERT(nvarchar,GETDATE(),111)=CONVERT(nvarchar,CREATE_DATE,111))
                BEGIN
                set @changetreasureid = 0
                set @changetreasureid = (
	                SELECT PIECE FROM USERS_PACKAGE(updlock)  WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID and TREASURE_ID = @TREASURE_ID 
                )
                set @changetreasureid2 = 0
                set @changetreasureid2 = (
	                SELECT PIECE FROM USERS_PACKAGE WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID and TREASURE_ID = @TREASURE_ID2 
                )
                if((@changetreasureid>= 1 AND  @changetreasureid2 >= 1 AND @TREASURE_ID <> @TREASURE_ID2 ) OR ( @changetreasureid >= 2 AND @TREASURE_ID = @TREASURE_ID2))
                begin
	                set @icti = @trId
	                if (@icti is not null)	
	                begin
	                insert into TREASURE_HUNT_LOG (ACTIVITY_ID,ACCOUNT_ID,READ_DOCUMENT_ID,TREASURE_ID,STATES,MESSAGE,CREATE_DATE) 
			                VALUES(@ACTIVITY_ID,@ACCOUNT_ID,0,@TREASURE_ID,@STATES,@MESSAGE,GETDATE())
                    insert into TREASURE_HUNT_LOG (ACTIVITY_ID,ACCOUNT_ID,READ_DOCUMENT_ID,TREASURE_ID,STATES,MESSAGE,CREATE_DATE) 
			                VALUES(@ACTIVITY_ID,@ACCOUNT_ID,0,@TREASURE_ID2,@STATES,@MESSAGE,GETDATE())
                    insert into TREASURE_HUNT_LOG (ACTIVITY_ID,ACCOUNT_ID,READ_DOCUMENT_ID,TREASURE_ID,STATES,MESSAGE,CREATE_DATE) 
			                VALUES(@ACTIVITY_ID,@ACCOUNT_ID,0,@icti,@STATES2,@MESSAGE2,GETDATE())
                    UPDATE USERS_PACKAGE SET PIECE = PIECE - 1 WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID and TREASURE_ID = @TREASURE_ID 
                    UPDATE USERS_PACKAGE SET PIECE = PIECE - 1 WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID and TREASURE_ID = @TREASURE_ID2 
                    UPDATE USERS_PACKAGE SET PIECE = PIECE + 1 WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID and TREASURE_ID = @icti 
	                select @icti as tid
	                end
	                else begin
	                select -1 as tid
	                end
                end
                else begin
	                select 0 as tid
                end
                end
                ELSE BEGIN
                   select -2 as tid 
                end
                COMMIT
            ";
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        var dt = SqlHelper.GetDataTable("TreasureHuntConnString", sql,
                    DbProviderFactories.CreateParameter("TreasureHuntConnString", "@ACTIVITY_ID", "@ACTIVITY_ID", actId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@ACCOUNT_ID", "@ACCOUNT_ID", userId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@TREASURE_ID", "@TREASURE_ID", treasureId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@TREASURE_ID2", "@TREASURE_ID2", treasureId2),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@trId", "@trId", trId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@MESSAGE", "@MESSAGE", Enum.GetName(typeof(TreasureHuntState), (int)TreasureHuntState.Chnage_Treasure)),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@STATES", "@STATES", (int)TreasureHuntState.Chnage_Treasure),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@STATES2", "@STATES2", (int)TreasureHuntState.Get_Treasure_by_Change),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@MESSAGE2", "@MESSAGE2", Enum.GetName(typeof(TreasureHuntState), (int)TreasureHuntState.Get_Treasure_by_Change)));
        int tid = int.Parse(dt.Rows[0]["tid"].ToString());
        log4.Error(userId + " change: " + treasureId.ToString() + " 2 pices to " + tid.ToString());
        return tid;
    }
    #endregion

    private int GetCheatRandomChangeTreasure(int userId, int actId, int treasureId, int treasureId2)
    {
        int getTreasureId = 0;
        Dictionary<int, int> usersTreasure = GetUsersTreasure(userId);

        int totalCount = 0;
        int limitSet = 99;
        int maxSuit = limitSet;
        int temp = 0;
        foreach (KeyValuePair<int, int> userItem in usersTreasure)
        {
            totalCount += userItem.Value;
            if (maxSuit > userItem.Value)
            {
                maxSuit = userItem.Value;
                temp = userItem.Key;
            }
        }

        if (totalCount >= (10 * (maxSuit + 1)))
        {
            getTreasureId = temp;
        }
        else
        {
            getTreasureId = RandomChangeTreasureId(userId,actId,treasureId,treasureId2);
        }
        log4.Error("totalCount:" + totalCount.ToString() + "__maxSuit:" + maxSuit + "__");
        return getTreasureId;
    }

    private int RandomChangeTreasureId(int userId, int actId, int treasureId, int treasureId2)
    {
        string sql = @"
                SELECT top 1 TREASURE_ID from treasure where TREASURE_ID <> @TREASURE_ID AND TREASURE_ID <> @TREASURE_ID2  and ACTIVITY_ID=@ACTIVITY_ID order by newid() 
            ";
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        var dt = SqlHelper.GetDataTable("TreasureHuntConnString", sql,
                    DbProviderFactories.CreateParameter("TreasureHuntConnString", "@ACTIVITY_ID", "@ACTIVITY_ID", actId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@TREASURE_ID", "@TREASURE_ID", treasureId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@TREASURE_ID2", "@TREASURE_ID2", treasureId2));
        int tid = int.Parse(dt.Rows[0]["TREASURE_ID"].ToString());
        return tid;
    }

    /// <summary>
    /// 判斷做弊模式是否開啟
    /// </summary>
    /// <returns></returns>
    private bool CheatModeOpen()
    {
        bool cheatIsopen = false;
        string sql = "IF ( GETDATE() BETWEEN (SELECT VALUE FROM dbo.CONFIG_DATA WHERE CONFIG_ID='CHEAT_MODE' AND ACTIVITY_ID = @ACTIVITY_ID) ";
        sql += "	AND (SELECT VALUE FROM dbo.CONFIG_DATA WHERE CONFIG_ID='CHEAT_MODE_END' AND ACTIVITY_ID = @ACTIVITY_ID)) ";
        sql += " begin ";
        sql += "  SELECT 1 ";
        sql += " END ";
        sql += " ELSE ";
        sql += " BEGIN ";
	    sql += "  SELECT 0 ";
        sql += " END ";
        string temp = GetSqlConnectionStringValue(sql);
        if (!string.IsNullOrEmpty(temp))
        {
            if (temp.CompareTo("1") == 0)
                cheatIsopen = true;
        }
        return cheatIsopen;
    }

    private string GetSqlConnectionStringValue(string sql)
    {
        string returnValue = "";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToString(o);
                    }
                }
                catch (Exception ex)
                {
                    ////throw ex;
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    public Treasure GetTreasureById(int id, int actId)
    {
        Treasure treasure = new Treasure();
        string sql = " SELECT * FROM TREASURE WHERE TREASURE_ID = @TREASURE_ID AND ACTIVITY_ID = @ACTIVITY_ID ";
        if (actId <= 0)
            actId = activity.Id;
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@TREASURE_ID", SqlDbType.Int).Value = id;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            treasure.TreasureId = Convert.ToInt32(dr["TREASURE_ID"].ToString());
                            treasure.TreasureName = dr["TREASURE_NAME"].ToString();
                            treasure.Icon = dr["ICON"].ToString();
                        }

                    }
                }
                catch (Exception ex)
                {
                    ////throw ex;
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
        return treasure;
    }

    /// <summary>
    /// 取得user缺少的寶物 備援程式
    /// </summary>
    /// <param name="actId"></param>
    /// <param name="userId"></param>
    /// <returns></returns>
    private int GetUserNeedsTreasure(int actId, int userId)
    {
        int treasureId = 0;
        Dictionary<int, int> usersTreasure = GetUsersTreasure( userId);

        int totalCount = 0;
        int limitSet = 99;
        int maxSuit = limitSet;
        int temp = 0;
        foreach (KeyValuePair<int, int> userItem in usersTreasure)
        {
            totalCount += userItem.Value;
            if (maxSuit > userItem.Value)
            {
                maxSuit = userItem.Value;
                temp = userItem.Key;
            }
        }

        if (totalCount >= (10 *  ( maxSuit + 1)))
        {
            treasureId = temp;
        }
        else
        {
            treasureId = GetTreasure();
        }
        log4.Error("totalCount:" + totalCount.ToString() + "__maxSuit:" + maxSuit +"__");
        return treasureId;
    }

    /// <summary>
    /// 隨機取寶
    /// </summary>
    /// <returns></returns>
    private int GetTreasure()
    {
        string sql = "select TOP(1) TREASURE_ID FROM TREASURE WHERE ACTIVITY_ID=@ACTIVITY_ID ORDER BY NEWID()";
        int treasureId =0;

        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        treasureId = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    ////throw ex;
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
        return treasureId;
    }

    /// <summary>
    /// 取得本次活動的寶物細項
    /// </summary>
    /// <returns></returns>
    public Dictionary<int, Treasure> GetTreasureDetal()
    {
        string sql=" SELECT * FROM TREASURE WHERE ACTIVITY_ID = @ACTIVITY_ID ";
        Dictionary<int, Treasure> tueasureSet = new Dictionary<int, Treasure>();
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            Treasure treasure = new Treasure();
                            treasure.TreasureId = Convert.ToInt32(dr["TREASURE_ID"]);
                            treasure.TreasureName = dr["TREASURE_NAME"].ToString();
                            treasure.Icon = dr["ICON"].ToString();
                            tueasureSet.Add(Convert.ToInt32(dr["TREASURE_ID"]), treasure);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
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
        return tueasureSet;
    }

    public Dictionary<int, int> GetUsersTreasure(int userAccountId)
    {
        Dictionary<int, int> usersTreasure = new Dictionary<int, int>();
        string sql = " SELECT * FROM USERS_PACKAGE WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID=@ACTIVITY_ID  ";
        Dictionary<string, Dictionary<string, string>> openWith = new Dictionary<string, Dictionary<string, string>>();

        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userAccountId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            usersTreasure.Add(Convert.ToInt32(dr["TREASURE_ID"]), Convert.ToInt32(dr["PIECE"]));
                        }
                    }
                }
                catch (Exception ex)
                {
                    ////throw ex;
                    log4.Error(ex.ToString());
                    return usersTreasure;
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


        return usersTreasure;
    }

    /// <summary>
    /// 判斷每日可取得的寶物上限是否超過
    /// </summary>
    /// <param name="accountId"></param>
    /// <param name="activityId"></param>
    /// <returns></returns>
    public bool CheckDailyGetLimit()
    {
        bool returnValue = false;
        string sql = " SELECT COUNT(STATES) FROM TREASURE_HUNT_LOG WHERE ACCOUNT_ID = @ACCOUNT_ID AND ACTIVITY_ID =@ACTIVITY_ID AND STATES = @STATES AND  ";
        sql += " MODIFY_DATE BETWEEN CONVERT(varchar(12) , getdate(), 111 ) AND CONVERT(varchar(12) , DateAdd(\"d\",1,getdate()), 111 ) ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = accountId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                cmd.Parameters.Add("@STATES", SqlDbType.Int).Value = (int)TreasureHuntState.Get_Treasure;
                try
                {
                    int dailyLimit = daily_set;
                    int dailyGet = 0;
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        dailyGet = Convert.ToInt32(o);
                        if (dailyLimit <= dailyGet)
                            returnValue = true;
                    }
                    //using (SqlDataReader dr = cmd.ExecuteReader())
                    //{
                    //    if (dr.Read())
                    //    {
                    //        returnValue = true;
                    //    }
                    //}
                }
                catch (Exception ex)
                {
                    //throw ex;
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
        return returnValue;
    }

    public IList GetTop(int pageIndex, int pageSize, int actId, string memberId, int treasureId, int totalSuits, int totalsuitUpper, double probability, int totaletsLowerBound, int totaletsUpperBound)
    {
       return GetTopAndQuery( pageIndex,  pageSize,  actId,  memberId,  treasureId,  totalSuits,  totalsuitUpper,  probability,  totaletsLowerBound,  totaletsUpperBound,"","",false,false);
    }

    /// <summary>
    /// 搜尋活動排名
    /// </summary>
    /// <param name="pageIndex"></param>
    /// <param name="pageSize"></param>
    /// <param name="actId">活動id</param>
    /// <param name="memberId">搜尋user</param>
    /// <param name="treasureId">搜尋寶物id</param>
    /// <param name="totalSuits">搜尋蒐集套數>=totalSuits</param>
    /// <returns></returns>
    public IList GetTopAndQuery(int pageIndex, int pageSize, int actId, string memberId, int treasureId, int totalSuits, int totalsuitUpper, double probability, int totaletsLowerBound, int totaletsUpperBound, string startTime, string endTime, bool reverse,bool onlyGetBox)
    { 
        IList topList = new ArrayList();
        string sql = " WITH subjectOrder AS ( SELECT LOGIN_ID,REALNAME,ACCOUNT.ACCOUNT_ID,NICKNAME,TOTAL_SUIT AS SUIT,CASE  WHEN TOTAL_SUIT > 0 THEN 'V' ELSE 'X' END \"QUALIFICATIONS\" ";
        sql += "  ,CASE  WHEN TOTAL_SUIT > @LIMIT_SET THEN @LIMIT_SET ELSE TOTAL_SUIT END \"LOTTERY\",CASE WHEN READCOUND IS NULL THEN 0 ELSE READCOUND END \"READCOUND\"  ";
        sql += "  ,COALESCE(TOTAL_TREASURE,0) AS \"TOTAL_TREASURE\" , COALESCE(CHANGE_TREASURE,0) AS \"CHANGE_TREASURE\",COALESCE(REAL_PACKAGE,0) AS \"REAL_PACKAGE\" ,ROW_NUMBER() OVER (ORDER BY TOTAL_SUIT DESC,REAL_PACKAGE DESC,GAME_TOP.ACCOUNT_ID) AS SN  ";
        sql += "   FROM ( ";
        sql += "      SELECT A.*, COALESCE(B.TOTAL_SUIT,0 ) AS \"TOTAL_SUIT\"  FROM (SELECT ACCOUNT_ID FROM ACCOUNT WHERE ACCOUNT_ID {3} IN ";
        sql += "   (SELECT DISTINCT ACCOUNT_ID FROM TREASURE_HUNT_LOG WHERE ACTIVITY_ID=@ACTIVITY_ID {0} ))  A  ";
        sql += "      LEFT JOIN ( ";
        sql += "        SELECT GAME_SUIT.ACCOUNT_ID, ";
        sql += "        CASE  WHEN TOTAL_KIND = @TOTAL_KIND THEN MIN_COUNT ELSE 0 END \"TOTAL_SUIT\" ";
        sql += "           FROM ( ";
        sql += "            SELECT ACCOUNT_ID,COUNT(TREASURE_ID) AS TOTAL_KIND,MIN(TREASURE_COUNT) AS MIN_COUNT FROM ";
        sql += "           ( SELECT ACCOUNT_ID,TREASURE_ID,PIECE AS TREASURE_COUNT from USERS_PACKAGE where ACTIVITY_ID=@ACTIVITY_ID ";
        sql += "                )AS GAME_DATA GROUP BY GAME_DATA.ACCOUNT_ID) AS GAME_SUIT ) B ON A.ACCOUNT_ID=B.ACCOUNT_ID  ";
        sql += "   ) AS GAME_TOP ";
        sql += "                LEFT JOIN ACCOUNT ON ACCOUNT.ACCOUNT_ID = GAME_TOP.ACCOUNT_ID   ";
        sql += "                LEFT JOIN (SELECT ACCOUNT_ID,COUNT(ACCOUNT_ID) AS READCOUND FROM TREASURE_HUNT_LOG ";
        sql += "                WHERE ACTIVITY_ID=@ACTIVITY_ID AND STATES <> 3 AND STATES <> 4 {2} GROUP BY ACCOUNT_ID) AS d ON d.ACCOUNT_ID =  dbo.ACCOUNT.ACCOUNT_ID   ";
        sql += "LEFT JOIN (SELECT ACCOUNT_ID,SUM(PIECE) AS REAL_PACKAGE FROM  USERS_PACKAGE WHERE ACTIVITY_ID=@ACTIVITY_ID  GROUP BY ACCOUNT_ID )AS PACKAGE ON PACKAGE.ACCOUNT_ID = ACCOUNT.ACCOUNT_ID";
        sql += "                LEFT JOIN (SELECT FFFF.ACCOUNT_ID,ISNULL( BBBBB.TOTAL_TREASURE,0) AS TOTAL_TREASURE,ISNULL( CCCCC.CHANGE_TREASURE,0) AS CHANGE_TREASURE ";
        sql += "  FROM ACCOUNT FFFF ";
        //sql += "  LEFT JOIN ( SELECT SUM(PIECE) AS TOTAL_TREASURE,account_id FROM USERS_PACKAGE  WHERE ACTIVITY_ID=@ACTIVITY_ID group by account_id) BBBBB ";
        sql += "  LEFT JOIN ( SELECT ACCOUNT_ID,COUNT(ACCOUNT_ID) AS TOTAL_TREASURE FROM TREASURE_HUNT_LOG ";
        sql += "  WHERE  ACTIVITY_ID=@ACTIVITY_ID AND STATES = @STATES {2}  GROUP BY ACCOUNT_ID) BBBBB ";
        sql += "  on FFFF.account_id=BBBBB.account_id ";
        sql += "  LEFT JOIN ( SELECT ACCOUNT_ID,COUNT(ACCOUNT_ID) AS CHANGE_TREASURE FROM TREASURE_HUNT_LOG ";
        sql += "   WHERE  ACTIVITY_ID=@ACTIVITY_ID AND STATES = 4 {2}  GROUP BY ACCOUNT_ID)  CCCCC ";
        sql += "   ON  CCCCC.account_id=FFFF.account_id";
        sql += "  ) AS F ON F.ACCOUNT_ID =  ACCOUNT.ACCOUNT_ID {1} ";
        
        sql += "  )";
        sql+= "SELECT * FROM subjectOrder WHERE SN BETWEEN @PAGEINDEX and @PAGESIZE ;";

        string queryMember = "";
        string queryTreasure = "";
        string dateString = "";
        if (probability >= 0)
        {
            queryMember += " LEFT JOIN (SELECT ACCOUNT_ID,ROUND(COUNT(TREASURE_ID)*1.00 / COUNT(ACCOUNT_ID),2) AS probability  FROM ";
            queryMember += "	TREASURE_HUNT_LOG  WHERE ACTIVITY_ID=@ACTIVITY_ID AND STATES <> 3 GROUP BY ACCOUNT_ID ";
            queryMember += "                )AS p ON p.ACCOUNT_ID = dbo.ACCOUNT.ACCOUNT_ID WHERE p.PROBABILITY < @PROBABILITY ";
            if (!string.IsNullOrEmpty(memberId))
            {
                queryMember += " AND ( LOGIN_ID LIKE @MEMBERID OR REALNAME LIKE @MEMBERID OR NICKNAME LIKE @MEMBERID ) ";
                if (totalSuits >= 0 && totalsuitUpper >= 0)
                {
                    queryMember += " AND TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER ";
                }
            }
            else if (totalSuits >= 0 && totalsuitUpper >=0)
            {
                queryMember += " AND TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER  ";
            }
        }
        else if (!string.IsNullOrEmpty(memberId))
        {
            queryMember += " WHERE ( LOGIN_ID LIKE @MEMBERID OR REALNAME LIKE @MEMBERID OR NICKNAME LIKE @MEMBERID ) ";
            if (totalSuits >= 0 && totalsuitUpper >= 0)
            {
                queryMember += " AND TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER ";
            }
        }
        else if (totalSuits >= 0 && totalsuitUpper >= 0)
        {
            queryMember += " WHERE TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER  ";
        }
        if(totaletsLowerBound >=0 && totaletsUpperBound >0)
        {
            if (string.IsNullOrEmpty(queryMember))
            {
                queryMember += " WHERE TOTAL_TREASURE >= @TOTALETSLOWERBOUND and TOTAL_TREASURE <= @TOTALETSUPPERBOUND ";
            }
            else
            {
                queryMember += " AND TOTAL_TREASURE >= @TOTALETSLOWERBOUND and TOTAL_TREASURE <= @TOTALETSUPPERBOUND ";
            }
        }
        if (treasureId > 0)
        {
            queryTreasure = " AND TREASURE_ID = @TREASUREID  ";
        }
        if (onlyGetBox)
        {
            queryTreasure += " AND STATES =" + (int)TreasureHuntState.Get_Box + "  ";
        }
        if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
        {
            dateString = " AND MODIFY_DATE BETWEEN @STARTTIME AND @ENDTIME ";
        }
        //反向搜尋用
        string reverseString = "";
        if (reverse)
            reverseString = " NOT ";
        sql = string.Format(sql, queryTreasure, queryMember, dateString, reverseString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@LIMIT_SET", SqlDbType.Int).Value = limit_set;
                if (actId == 0)
                    actId = activity.Id;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                cmd.Parameters.Add("@STATES", SqlDbType.Int).Value = (int)TreasureHuntState.Get_Treasure;
                cmd.Parameters.Add("@PAGEINDEX", SqlDbType.Int).Value = (1 + (pageIndex-1) * pageSize); 
                cmd.Parameters.Add("@PAGESIZE", SqlDbType.Int).Value = (pageIndex) * pageSize;
                cmd.Parameters.Add("@MEMBERID", SqlDbType.NVarChar, 50).Value = '%'+memberId+'%';
                cmd.Parameters.Add("@TREASUREID", SqlDbType.Int).Value = treasureId;
                cmd.Parameters.Add("@TOTALSUITS", SqlDbType.Int).Value = totalSuits;
                cmd.Parameters.Add("@TOTALSUITUPPER", SqlDbType.Int).Value = totalsuitUpper;
                cmd.Parameters.Add("@PROBABILITY", SqlDbType.Float).Value = probability;
                cmd.Parameters.Add("@TOTAL_KIND", SqlDbType.Float).Value = GetTotalTreasureCount(actId);
                cmd.Parameters.Add("@TOTALETSLOWERBOUND", SqlDbType.Int).Value = totaletsLowerBound;
                cmd.Parameters.Add("@TOTALETSUPPERBOUND", SqlDbType.Int).Value = totaletsUpperBound;
                if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
                {
                    cmd.Parameters.Add("@STARTTIME", SqlDbType.DateTime).Value = DateTime.ParseExact(startTime, "yyyy/MM/dd", null);
                    cmd.Parameters.Add("@ENDTIME", SqlDbType.DateTime).Value = DateTime.ParseExact(endTime, "yyyy/MM/dd", null).AddDays(1);
                }
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            Treasure_Top treasure_top = new Treasure_Top();
                            treasure_top.AccountId = Convert.ToInt32(dr["ACCOUNT_ID"].ToString());
                            treasure_top.LoginId = dr["LOGIN_ID"].ToString();
                            treasure_top.RealName = dr["REALNAME"].ToString();
                            treasure_top.NickName = dr["NICKNAME"].ToString();
                            treasure_top.TotalSuit =Convert.ToInt32(dr["SUIT"].ToString());
                            treasure_top.Qualifications = dr["QUALIFICATIONS"].ToString();
                            treasure_top.Lottery = Convert.ToInt32(dr["LOTTERY"].ToString());
                            treasure_top.TotalReadCount = dr["READCOUND"].ToString();
                            treasure_top.Total_Treasure = Convert.ToInt32(dr["TOTAL_TREASURE"].ToString());
                            treasure_top.Change_Treasure = Convert.ToInt32(dr["CHANGE_TREASURE"].ToString());
                            treasure_top.Real_Treasure = Convert.ToInt32(dr["REAL_PACKAGE"].ToString());  
                            topList.Add(treasure_top);
                        }
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
                    log4.Error(sql);
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
        return topList;
    }


    public int GetTotalGameMunberTop(int actId, string memberId, int treasureId, int totalSuits, int totalsuitUpper, double probability, int totaletsLowerBound, int totaletsUpperBound, string startTime, string endTime)
    {
        return GetTotalGameMunberTopImpl( actId, memberId,treasureId, totalSuits,totalsuitUpper,probability,totaletsLowerBound,totaletsUpperBound,startTime,endTime,false);
    }

    /// <summary>
    /// 查詢用的總比數
    /// </summary>
    /// <param name="actId"></param>
    /// <returns></returns>
    public int GetTotalGameMunberTopImpl(int actId, string memberId, int treasureId, int totalSuits, int totalsuitUpper, double probability, int totaletsLowerBound, int totaletsUpperBound,string startTime,string endTime,bool onlyGetBox)
    {
        int returnValue = 0;
        int act_id = 1;
        if (actId <= 0)
        {
            act_id = activity.Id;
        }
        else
        {
            act_id = actId;
        }

        IList topList = new ArrayList();
        string sql = " SELECT COUNT(A.ACCOUNT_ID) FROM (SELECT DISTINCT ACCOUNT_ID FROM TREASURE_HUNT_LOG WHERE ACTIVITY_ID=@ACTIVITY_ID {0} )  A  ";
        sql += "      LEFT JOIN ( ";
        sql += "        SELECT GAME_SUIT.ACCOUNT_ID, ";
        sql += "        CASE  WHEN TOTAL_KIND = @TOTAL_KIND THEN MIN_COUNT ELSE 0 END \"TOTAL_SUIT\" ";
        sql += "           FROM ( ";
        sql += "            SELECT ACCOUNT_ID,COUNT(TREASURE_ID) AS TOTAL_KIND,MIN(TREASURE_COUNT) AS MIN_COUNT FROM ";
        sql += "           (  SELECT ACCOUNT_ID,TREASURE_ID,PIECE AS TREASURE_COUNT from USERS_PACKAGE where ACTIVITY_ID=@ACTIVITY_ID  ";
        sql += "               )AS GAME_DATA GROUP BY GAME_DATA.ACCOUNT_ID) AS GAME_SUIT ) B ON A.ACCOUNT_ID=B.ACCOUNT_ID  ";
       
        sql += "                LEFT JOIN ACCOUNT ON ACCOUNT.ACCOUNT_ID = A.ACCOUNT_ID   ";
        sql += "                LEFT JOIN (SELECT ACCOUNT_ID,COUNT(ACCOUNT_ID) AS READCOUND FROM dbo.TREASURE_HUNT_LOG ";
        sql += "                WHERE ACTIVITY_ID=@ACTIVITY_ID AND  STATES <> 3 {2}  GROUP BY ACCOUNT_ID) AS d ON d.ACCOUNT_ID =  dbo.ACCOUNT.ACCOUNT_ID   ";
        sql += "                LEFT JOIN (SELECT FFFF.ACCOUNT_ID,ISNULL( BBBBB.TOTAL_TREASURE,0) AS TOTAL_TREASURE ";
        sql += "  FROM ACCOUNT FFFF ";
        sql += "  LEFT JOIN ( SELECT ACCOUNT_ID,COUNT(ACCOUNT_ID) AS TOTAL_TREASURE FROM dbo.TREASURE_HUNT_LOG ";
        sql += "  WHERE  ACTIVITY_ID=@ACTIVITY_ID AND STATES = @STATES {2}   GROUP BY ACCOUNT_ID) BBBBB ";
        sql += "  on FFFF.account_id=BBBBB.account_id ) AS F ON F.ACCOUNT_ID =  dbo.ACCOUNT.ACCOUNT_ID {1} ";

        string queryMember = "";
        string queryTreasure = "";
        if (probability >= 0)
        {
            queryMember += " LEFT JOIN (SELECT ACCOUNT_ID,ROUND(COUNT(TREASURE_ID)*1.00 / COUNT(ACCOUNT_ID),2) AS probability  FROM ";
            queryMember += "	TREASURE_HUNT_LOG  WHERE ACTIVITY_ID=@ACTIVITY_ID AND STATES <> 3 GROUP BY ACCOUNT_ID ";
            queryMember += "                )AS p ON p.ACCOUNT_ID = dbo.ACCOUNT.ACCOUNT_ID WHERE p.PROBABILITY < @PROBABILITY ";
            if (!string.IsNullOrEmpty(memberId))
            {
                queryMember += " AND ( LOGIN_ID LIKE @MEMBERID OR REALNAME LIKE @MEMBERID OR NICKNAME LIKE @MEMBERID ) ";
                if (totalSuits >= 0 && totalsuitUpper >= 0)
                {
                    queryMember += " AND TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER ";
                }
            }
            else if (totalSuits >= 0 && totalsuitUpper >= 0)
            {
                queryMember += " AND TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER  ";
            }


        }
        else if (!string.IsNullOrEmpty(memberId))
        {
            queryMember += " WHERE ( LOGIN_ID LIKE @MEMBERID OR REALNAME LIKE @MEMBERID OR NICKNAME LIKE @MEMBERID ) ";
            if (totalSuits >= 0 && totalsuitUpper >= 0)
            {
                queryMember += " AND TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER ";
            }
        }
        else if (totalSuits >= 0 && totalsuitUpper >= 0)
        {
            queryMember += " WHERE TOTAL_SUIT BETWEEN @TOTALSUITS AND @TOTALSUITUPPER  ";
        }
        if (totaletsLowerBound >= 0 && totaletsUpperBound > 0)
        {
            if (string.IsNullOrEmpty(queryMember))
            {
                queryMember += " WHERE TOTAL_TREASURE >= @TOTALETSLOWERBOUND AND TOTAL_TREASURE <= @TOTALETSUPPERBOUND ";
            }
            else
            {
                queryMember += " AND TOTAL_TREASURE >= @TOTALETSLOWERBOUND AND TOTAL_TREASURE <= @TOTALETSUPPERBOUND ";
            }
        }
        if (treasureId > 0)
        {
            queryTreasure = " AND TREASURE_ID = @TREASUREID  ";
        }
        if (onlyGetBox)
        {
            queryTreasure += " AND STATES =" + (int)TreasureHuntState.Get_Box + "  ";
        }
        string dateString = "";
        if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
        {
            dateString = " AND MODIFY_DATE BETWEEN @STARTTIME AND @ENDTIME ";
        }
        sql = string.Format(sql, queryTreasure, queryMember, dateString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@LIMIT_SET", SqlDbType.Int).Value = limit_set;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = act_id;
                cmd.Parameters.Add("@STATES", SqlDbType.Int).Value = (int)TreasureHuntState.Get_Treasure;
                cmd.Parameters.Add("@MEMBERID", SqlDbType.NVarChar, 50).Value = '%' + memberId + '%';
                cmd.Parameters.Add("@TREASUREID", SqlDbType.Int).Value = treasureId;
                cmd.Parameters.Add("@TOTALSUITS", SqlDbType.Int).Value = totalSuits;
                cmd.Parameters.Add("@TOTALSUITUPPER", SqlDbType.Int).Value = totalsuitUpper;
                cmd.Parameters.Add("@PROBABILITY", SqlDbType.Float).Value = probability;
                cmd.Parameters.Add("@TOTAL_KIND", SqlDbType.Float).Value = GetTotalTreasureCount(actId);
                cmd.Parameters.Add("@TOTALETSLOWERBOUND", SqlDbType.Int).Value = totaletsLowerBound;
                cmd.Parameters.Add("@TOTALETSUPPERBOUND", SqlDbType.Int).Value = totaletsUpperBound;
                if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
                {
                    cmd.Parameters.Add("@STARTTIME", SqlDbType.DateTime).Value = DateTime.ParseExact(startTime, "yyyy/MM/dd", null);
                    cmd.Parameters.Add("@ENDTIME", SqlDbType.DateTime).Value = DateTime.ParseExact(endTime, "yyyy/MM/dd", null).AddDays(1);
                }
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
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
        return returnValue;
    }

    /// <summary>
    /// 取得該次活動所有有分數的總人數 傳入0以下代表本次活動
    /// </summary>
    /// <param name="actId"></param>
    /// <returns></returns>
    public int GetTotalGameMunber(int actId)
    {
        int returnValue=0;
        int act_id = 1;
        if (actId <= 0)
        {
            act_id = activity.Id;
        }
        else
        {
            act_id = actId;
        }


        string sql = " select COUNT(D.ACCOUNT) from ( ";
        sql += "SELECT COUNT(DISTINCT ACCOUNT_ID) as ACCOUNT  FROM dbo.TREASURE_HUNT_LOG WHERE ACTIVITY_ID = @ACTIVITY_ID GROUP BY ACCOUNT_ID ";
        sql += ") as D";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = act_id;
                cmd.Parameters.Add("@STATES", SqlDbType.Int).Value = (int)TreasureHuntState.Get_Treasure;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
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
        return returnValue;
    }

    public int GetUsersLottery(string id)
    {
        loginId = id.Trim();
        return GetUsersLottery(accountId);
    }

    public int GetUsersLottery(int id)
    {
        if (id <= 0)
            id = accountId;
        int returnValue = 0;
        string sql = " SELECT CASE  WHEN TOTAL_SUIT > @LIMIT_SET THEN @LIMIT_SET ELSE TOTAL_SUIT END \"LOTTERY\" FROM ";
        sql += "    ( ";
        sql += "        SELECT  GAME_SUIT.ACCOUNT_ID, ";
        sql += "        CASE  WHEN TOTAL_KIND = @TREASURE_COUNT  THEN MIN_COUNT ELSE 0 END \"TOTAL_SUIT\" ";
        sql += "        FROM ( ";
        sql += "        SELECT ACCOUNT_ID,COUNT(TREASURE_ID) AS TOTAL_KIND,MIN(TREASURE_COUNT) AS MIN_COUNT FROM ";
        sql += "        ( SELECT ACCOUNT_ID,TREASURE_ID,PIECE AS TREASURE_COUNT  FROM USERS_PACKAGE WHERE ACTIVITY_ID=@ACTIVITY_ID AND  ACCOUNT_ID = @ACCOUNT_ID  ";
        sql += "           )AS GAME_DATA GROUP BY GAME_DATA.ACCOUNT_ID) AS GAME_SUIT )AS Game_top ";
        sql += "        WHERE ACCOUNT_ID =@ACCOUNT_ID ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = id;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                cmd.Parameters.Add("@LIMIT_SET", SqlDbType.Int).Value = limit_set;
                cmd.Parameters.Add("@TREASURE_COUNT ", SqlDbType.Int).Value = GetTotalTreasureCount(activity.Id);
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
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
        return returnValue;
    }

    public bool CanChangeTreasure(int userId,int actId)
    {
        string sql = @"
                SELECT * FROM  TREASURE_HUNT_LOG 
                WHERE ACCOUNT_ID = @ACCOUNT_ID AND STATES=@STATES AND ACTIVITY_ID=@ACTIVITY_ID 
                 AND CONVERT(nvarchar,GETDATE(),111)=CONVERT(nvarchar,CREATE_DATE,111)
        ";
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        var dt = SqlHelper.GetDataTable("TreasureHuntConnString", sql,
                    DbProviderFactories.CreateParameter("TreasureHuntConnString", "@ACTIVITY_ID", "@ACTIVITY_ID", actId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@ACCOUNT_ID", "@ACCOUNT_ID", userId),
                     DbProviderFactories.CreateParameter("TreasureHuntConnString", "@STATES", "@STATES", (int)TreasureHuntState.Get_Treasure_by_Change));
        if(dt.Rows.Count >0)
            return false;
        return true;
    }

    public int GetUsersTotalSuit(string id)
    {
        loginId = id.Trim();
        return GetUsersTotalSuit(accountId);
    }

    /// <summary>
    /// 取得user的總套數
    /// </summary>
    /// <param name="id"></param>
    /// <returns></returns>
    public int GetUsersTotalSuit(int id)
    {
        if (id <= 0)
            id = accountId;
        int returnValue = 0;
        string sql = " SELECT  TOTAL_SUIT  FROM ";
        sql += "    ( ";
        sql += "        SELECT  GAME_SUIT.ACCOUNT_ID, ";
        sql += "        CASE  WHEN TOTAL_KIND = @TREASURE_COUNT  THEN MIN_COUNT ELSE 0 END \"TOTAL_SUIT\" ";
        sql += "        FROM ( ";
        sql += "        SELECT ACCOUNT_ID,COUNT(TREASURE_ID) AS TOTAL_KIND,MIN(TREASURE_COUNT) AS MIN_COUNT FROM ";
        sql += "        ( SELECT ACCOUNT_ID,TREASURE_ID,PIECE AS TREASURE_COUNT  FROM USERS_PACKAGE WHERE ACTIVITY_ID=@ACTIVITY_ID AND  ACCOUNT_ID = @ACCOUNT_ID   ";
        sql += "           )AS GAME_DATA GROUP BY GAME_DATA.ACCOUNT_ID) AS GAME_SUIT )AS Game_top ";
        sql += "        WHERE ACCOUNT_ID =@ACCOUNT_ID ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = id;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activity.Id;
                cmd.Parameters.Add("@LIMIT_SET", SqlDbType.Int).Value = limit_set;
                cmd.Parameters.Add("@TREASURE_COUNT ", SqlDbType.Int).Value = GetTotalTreasureCount(activity.Id);
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
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
        return returnValue;
    }


    /// <summary>
    /// 取得user的寶物種類數量
    /// </summary>
    /// <param name="actId"></param>
    /// <param name="userId"></param>
    /// <param name="startTime"></param>
    /// <param name="endTime"></param>
    /// <returns></returns>
    public Dictionary<int, int> GetUsetAllKindTreasureCount(int actId, int userId, string startTime, string endTime)
    {
        Dictionary<int, int> returnValue = new Dictionary<int, int>();
        string sql = " SELECT TREASURE_ID,COUNT(TREASURE_ID) AS TOTAL_COUNT FROM dbo.TREASURE_HUNT_LOG  ";
        sql += "            WHERE TREASURE_ID IS NOT NULL AND STATES <> 3 AND STATES <> 4 AND ACCOUNT_ID = @ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID  ";
	    sql += "            {0} ";
        sql += "       GROUP BY TREASURE_ID ";
        string dateString = "";
        if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
        {
            dateString = " AND MODIFY_DATE BETWEEN @STARTTIME AND @ENDTIME ";
        }
        sql = string.Format(sql,dateString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userId;
                if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
                {
                    cmd.Parameters.Add("@STARTTIME", SqlDbType.DateTime).Value = DateTime.ParseExact(startTime, "yyyy/MM/dd", null);
                    cmd.Parameters.Add("@ENDTIME", SqlDbType.DateTime).Value = DateTime.ParseExact(endTime, "yyyy/MM/dd", null).AddDays(1);
                }
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            if (!(dr["TREASURE_ID"] is DBNull) && !(dr["TOTAL_COUNT"] is DBNull))
                                returnValue.Add(Convert.ToInt32(dr["TREASURE_ID"].ToString()), Convert.ToInt32(dr["TOTAL_COUNT"].ToString()));
                        }

                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    /// <summary>
    /// 取得所有寶物總類數量
    /// </summary>
    /// <param name="actId"></param>
    /// <returns></returns>
    private int GetTotalTreasureCount(int actId)
    {
        int returnValue = 0;
        if (actId <= 0) actId = activity.Id;
        string sql = " SELECT COUNT(DISTINCT TREASURE.TREASURE_ID) FROM TREASURE WHERE ACTIVITY_ID=@ACTIVITY_ID   ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    //throw ex;
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
        return returnValue;
    }
    /// <summary>
    /// 取得當次個別寶物總發寶數(由尋寶獲得)
    /// </summary>
    /// <param name="actId"></param>
    /// <returns></returns>
    public Dictionary<int, int> GetAllFindedTreasureCount(int actId)
    {
        Dictionary<int, int> returnValue = new Dictionary<int, int>();
        string sql = " SELECT COUNT(TREASURE_ID) AS TOTAL_COUNT,TREASURE_ID FROM TREASURE_HUNT_LOG ";
        sql += " WHERE TREASURE_ID IS NOT NULL AND STATES <> 3 AND STATES <> 4 AND ACTIVITY_ID=@ACTIVITY_ID GROUP BY TREASURE_ID";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while(dr.Read())
                        {
                            if (!(dr["TREASURE_ID"] is DBNull) && !(dr["TOTAL_COUNT"] is DBNull))
                            returnValue.Add(Convert.ToInt32(dr["TREASURE_ID"].ToString()), Convert.ToInt32(dr["TOTAL_COUNT"].ToString()));
                        }

                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    /// <summary>
    /// 取得當次活動個別寶物總個數
    /// </summary>
    /// <param name="actId"></param>
    /// <returns></returns>
    public Dictionary<int, int> GetAllUserBackageTreasureCount(int actId)
    {
        Dictionary<int, int> returnValue = new Dictionary<int, int>();
        string sql = @"SELECT TREASURE_ID,SUM(PIECE) as TOTAL_COUNT FROM USERS_PACKAGE WHERE ACTIVITY_ID = @ACTIVITY_ID
                        GROUP BY TREASURE_ID";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            if (!(dr["TREASURE_ID"] is DBNull) && !(dr["TOTAL_COUNT"] is DBNull))
                                returnValue.Add(Convert.ToInt32(dr["TREASURE_ID"].ToString()), Convert.ToInt32(dr["TOTAL_COUNT"].ToString()));
                        }

                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    //找出文章id
    public int getDocumentId(string url)
    {
        int id = -1;
        string temp = url.Split('?')[1];
        string[] param = temp.Split('&');
        string type = "";
        type = GetDocumentIdTypeName(url);
        try
        {
            foreach (string s in param)
            {
                if (s.Contains(type))
                {
                    id = Convert.ToInt32(s.Substring(s.IndexOf('=') + 1));
                }
            }
        }
        catch (Exception ex)
        {
            log4.Error("文章id錯誤 url=" + url + "\n accountId=" + account_id + "\n" + ex.ToString());
            return id;
        }
        return id;
    }
    /// <summary>
    /// 取得文章type
    /// </summary>
    /// <param name="url"></param>
    /// <returns></returns>
    public string GetDocumentIdTypeName(string url)
    {
        string type = "xitem";
        if (url.Contains("subject"))
        {
            type = "xitem";
        }
        else if (url.Contains("knowledge"))
        {
            type = "articleid";
        }
        else if (url.Contains("category"))
        {
            type = "reportid";
        }
        else if (url.Contains("jigsaw2010"))
        {
            type = "item";
        }
        return type;
    }
    /// <summary>
    /// 取得文章type&Name
    /// </summary>
    /// <param name="url"></param>
    /// <returns></returns>
    public string GetDocumentIdAndType(string url)
    {
        string type = GetDocumentIdTypeName(url);
        string documenttypeandname = "";
        string temp = url.Split('?')[1];
        string[] param = temp.Split('&');
        foreach (string s in param)
        {
            if (s.Contains(type))
            {
                documenttypeandname = s;
                break;
            }
        }
        return documenttypeandname;
    }
    /// <summary>
    /// 取得單元名稱
    /// </summary>
    /// <param name="url"></param>
    /// <returns></returns>
    public string GetActivityUnitName(string url)
    {
        string returnName = "";
        int id = -1;
        string temp = url.Split('?')[1];
        string[] param = temp.Split('&');
        string type = "";
        if (url.Contains("subject"))
        {
            type = "mp";
        }
        else if (url.Contains("knowledge"))
        {
            returnName = "知識家";
            return returnName;
        }
        else if (url.Contains("/category"))
        {
            returnName = "知識庫";
            return returnName;
        }
        else if (url.Contains("/jigsaw2010"))
        {
            returnName = "農漁生產地圖";
            return returnName;
        }
        else 
        {
            //ct.asp
            return "站內單元";
        }
        foreach (string s in param)
        {
            if (s.Contains(type))
            {
                id = Convert.ToInt32(s.Substring(s.IndexOf('=') + 1));
            }
        }
        //蝴蝶蘭 鳳梨用hard code
        if(id== 7)
            return "鳳梨主題館";
        if (id == 5)
            return "蝴蝶蘭主題館";
        if (id != -1)
        {
            string sql = "SELECT CTROOTNAME FROM CATTREEROOT WHERE CTROOTID = @CTROOTID";
            using (SqlConnection conn = TreasureHuntDB.createCoaConnection())
            {
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("@CTROOTID", SqlDbType.Int).Value = id;
                    try
                    {
                        conn.Open();
                        object o = cmd.ExecuteScalar();
                        if (!(o is DBNull))
                        {
                            returnName = Convert.ToString(o);
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
        else if (url.Contains("knowledge"))
        {
            returnName = "知識家文章";
        }
        return returnName;
    }

    /// <summary>
    /// 找尋文章title 傳入文章id
    /// </summary>
    /// <param name="iCuItem"></param>
    /// <returns></returns>
    public string  getDocumentTitle(int iCuItem)
    {
        //string sql = " SELECT  HTX.*, XR1.DEPTNAME, U.CTUNITNAME ";
        //sql+=" , (SELECT COUNT(*) FROM CUDTATTACH AS DHTX  ";
        //sql+=" WHERE BLIST='Y' AND DHTX.XICUITEM=HTX.ICUITEM) AS ATTACHCOUNT    ";
        //sql+=" , (SELECT COUNT(*) FROM CUDTPAGE AS PHTX";
       // sql+="  WHERE BLIST='Y' AND PHTX.XICUITEM=HTX.ICUITEM) AS PAGECOUNT ";
        //sql+="   FROM CUDTGENERIC AS HTX LEFT JOIN DEPT AS XR1 ON XR1.DEPTID=HTX.IDEPT ";
        //sql+="  LEFT JOIN CTUNIT AS U ON U.CTUNITID=HTX.ICTUNIT ";
        //sql+="  WHERE HTX.ICUITEM= @ICUITEM";
        string sql = " SELECT STITLE FROM CUDTGENERIC WHERE ICUITEM =@ICUITEM ";
        string returnValue="";
        using (SqlConnection conn = TreasureHuntDB.createCoaConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ICUITEM", SqlDbType.Int).Value = iCuItem;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToString(o);
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }


    /// <summary>
    /// 搜尋 尋寶紀錄
    /// </summary>
    /// <param name="userAcountId"></param>
    /// <param name="activityId"></param>
    /// <param name="states"></param>
    public IList GetSearchTreasureRecord(int userAcountId, int activityId, int states, int pageIndex, int pageSize)
    {
        IList treasureList = new ArrayList();
        string sql = "  WITH SUBJECTORDER AS (  ";
        sql += " SELECT DOCUMENT_TITLE,UNIT_NAME,URL,STATES,TREASURE_NAME,Convert(varchar(10),TREASURE_HUNT_LOG.CREATE_DATE,111)AS GET_DATE ";
        sql += " ,ROW_NUMBER() OVER (ORDER BY TREASURE_HUNT_LOG.CREATE_DATE) AS SN FROM dbo.TREASURE_HUNT_LOG ";
               sql += " LEFT JOIN dbo.READ_DOCUMENT_RECORD ON TREASURE_HUNT_LOG.READ_DOCUMENT_ID = READ_DOCUMENT_RECORD.READ_DOCUMENT_ID ";
               sql += " LEFT JOIN TREASURE ON TREASURE.TREASURE_ID = TREASURE_HUNT_LOG.TREASURE_ID AND TREASURE.ACTIVITY_ID = TREASURE_HUNT_LOG.ACTIVITY_ID ";
               sql += " WHERE TREASURE_HUNT_LOG.ACCOUNT_ID = @ACCOUNT_ID AND TREASURE_HUNT_LOG.ACTIVITY_ID =@ACTIVITY_ID  {0})";
               sql += " SELECT * FROM SUBJECTORDER WHERE SN BETWEEN @PAGEINDEX and @PAGESIZE ; ";
        string stateString = "";
        if (states >= 0 && states != 4 && states != 2)
        {
            stateString = " AND STATES = @STATES   ";
        }
        else if (states == -2)
        {
            stateString = " AND STATES <> 3 AND STATES<>4  ";
        }
        else if (states == 4)
        {
            stateString = " AND STATES in (3,4) ";
        }
        else if (states == 2)
        {
            stateString = " AND STATES in (2,3,4) ";
        }
       sql = string.Format(sql, stateString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userAcountId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activityId;
                cmd.Parameters.Add("@STATES", SqlDbType.Int).Value = states;
                cmd.Parameters.Add("@PAGEINDEX", SqlDbType.Int).Value = (1 + (pageIndex - 1) * pageSize);
                cmd.Parameters.Add("@PAGESIZE", SqlDbType.Int).Value = (pageIndex) * pageSize;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            TreasureLogDetal treasureLogDetal = new TreasureLogDetal();
                            treasureLogDetal.UnitName = dr["UNIT_NAME"].ToString();
                            treasureLogDetal.DocumentTitle = dr["DOCUMENT_TITLE"].ToString();
                            treasureLogDetal.Url = dr["URL"].ToString();
                            treasureLogDetal.GetDate = dr["GET_DATE"].ToString();
                            treasureLogDetal.States = Convert.ToInt32(dr["STATES"].ToString());
                            if (dr["TREASURE_NAME"] != null)
                            {
                                treasureLogDetal.TreasureName = dr["TREASURE_NAME"].ToString();
                            }
                            else
                            {
                                treasureLogDetal.TreasureName = "";
                            }
                            treasureList.Add(treasureLogDetal);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return treasureList;
    }

    public int getTreasureHuntLogCount(int userAcountId,int activityId,int states)
    {
        int returnValue = 0;
        string sql = "  SELECT COUNT(TREASURE_HUNT_LOG_ID) ";
               sql += " FROM dbo.TREASURE_HUNT_LOG ";
               sql += " LEFT JOIN dbo.READ_DOCUMENT_RECORD ON TREASURE_HUNT_LOG.READ_DOCUMENT_ID = READ_DOCUMENT_RECORD.READ_DOCUMENT_ID ";
               sql += " LEFT JOIN TREASURE ON TREASURE.TREASURE_ID = TREASURE_HUNT_LOG.TREASURE_ID AND TREASURE.ACTIVITY_ID = TREASURE_HUNT_LOG.ACTIVITY_ID ";
               sql += " WHERE TREASURE_HUNT_LOG.ACCOUNT_ID = @ACCOUNT_ID AND TREASURE_HUNT_LOG.ACTIVITY_ID =@ACTIVITY_ID {0} ";
               string stateString = "";
       if (states >= 0 && states != 4 && states != 2)
       {
           stateString = " AND STATES = @STATES   ";
       }
       else if (states == -2)
       {
           stateString = " AND STATES <> 3 AND STATES<>4  ";
       }
       else if (states == 4)
       {
           stateString = " AND STATES in (3,4) ";
       }
       else if (states == 2)
       {
           stateString = " AND STATES in (2,3,4) ";
       }
       sql = string.Format(sql, stateString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userAcountId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = activityId;
                cmd.Parameters.Add("@STATES", SqlDbType.Int).Value = states;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    public IList GetUserVotesLottery(int actId, int userId)
    {
        IList returnValue = new ArrayList();
        string sql = " SELECT GIFT_ID,VOTES FROM LOTTERY_VOTE WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID  ";
        string stateString = "";
        if (actId <= 0)
            actId = activity.Id;
        sql = string.Format(sql, stateString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            LotteryGift lotteryGift = new LotteryGift();
                            lotteryGift.GiftId = dr["GIFT_ID"].ToString();
                            lotteryGift.Votes = Convert.ToInt32(dr["VOTES"].ToString());
                            returnValue.Add(lotteryGift);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }


    public bool SetUserVotesLottery(int actId, int userId,string giftId)
    {
        bool returnValue= false;
        log4.Error(actId + "__" + userId + "_" + giftId);
        string sql = " SELECT VOTES FROM LOTTERY_VOTE WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID AND GIFT_ID = @GIFT_ID  ";
        string stateString = "";
        if (actId <= 0)
            actId = activity.Id;
        sql = string.Format(sql, stateString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                cmd.Parameters.Add("@GIFT_ID", SqlDbType.VarChar,50).Value = giftId;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull) && Convert.ToInt32(o) != 0)
                    {
                        int count =  Convert.ToInt32(o);
                        log4.Error(actId + "__" + userId + "_" + giftId + "__" + count);
                        using (SqlCommand cmd_vote = new SqlCommand())
                        {
                            sql = " UPDATE LOTTERY_VOTE SET VOTES=@VOTES WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID AND GIFT_ID = @GIFT_ID  ";
                            count++;
                            cmd_vote.CommandText = sql;
                            cmd_vote.Connection = conn;
                            cmd_vote.Parameters.Clear();
                            cmd_vote.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userId;
                            cmd_vote.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                            cmd_vote.Parameters.Add("@GIFT_ID", SqlDbType.VarChar, 50).Value = giftId;
                            cmd_vote.Parameters.Add("@VOTES", SqlDbType.Int).Value = count;
                            cmd_vote.ExecuteNonQuery();
                            returnValue = true;
                        }
                    }
                    else
                    {
                        int count = 1;
                        log4.Error(actId + "__" + userId + "_" + giftId + "__" + count);
                        using (SqlCommand cmd_vote = new SqlCommand())
                        {
                            sql = " INSERT INTO   LOTTERY_VOTE (ACCOUNT_ID,ACTIVITY_ID,GIFT_ID,VOTES) ";
                            sql += " VALUES(@ACCOUNT_ID,@ACTIVITY_ID,@GIFT_ID,@VOTES) ";
                            cmd_vote.CommandText = sql;
                            cmd_vote.Connection = conn;
                            cmd_vote.Parameters.Clear();
                            cmd_vote.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userId;
                            cmd_vote.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                            cmd_vote.Parameters.Add("@GIFT_ID", SqlDbType.VarChar, 50).Value = giftId;
                            cmd_vote.Parameters.Add("@VOTES", SqlDbType.Int).Value = count;
                            cmd_vote.ExecuteNonQuery();
                            returnValue = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    return false;
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
        return returnValue;
    }

    public int GetUserVotesLotteryCount(int actId, int userId)
    {
        int returnValue = 0;
        string sql = "  SELECT  SUM(VOTES) AS VOTES  FROM LOTTERY_VOTE WHERE ACCOUNT_ID =@ACCOUNT_ID AND ACTIVITY_ID = @ACTIVITY_ID ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        returnValue = Convert.ToInt32(o);
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    /// <summary>
    /// 判斷投套數是否開始
    /// </summary>
    /// <returns></returns>
    public bool VoteLotteryIsEnable()
    {
        bool returnValue = false;
        string temp = GetConfigData("LOTTERY_START");
        if(string.IsNullOrEmpty(temp))
            return false;

        DateTime startTime = Convert.ToDateTime(temp);
        temp = GetConfigData("LOTTERY_END");
        if (string.IsNullOrEmpty(temp))
            return false;
        DateTime endTime = Convert.ToDateTime(temp);
        string sql = "IF(GETDATE() BETWEEN @STARTTIME AND @ENDTIME) ";
        sql += " BEGIN ";
        sql += " SELECT 1 ";
        sql += " END ELSE ";
        sql += "  BEGIN ";
        sql += " SELECT 0 ";
        sql += " END ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@STARTTIME", SqlDbType.DateTime).Value = startTime;
                cmd.Parameters.Add("@ENDTIME", SqlDbType.DateTime).Value = endTime;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                        if (Convert.ToInt32(o)==1)
                            returnValue = true; ;
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    public IList GetAllActivityList()
    {
        IList returnValue = new ArrayList();
        string sql = " SELECT ACTIVITY_ID,ACTIVITY_NAME FROM ACTIVITY ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            Treasure tr = new Treasure();
                            tr.TreasureId = Convert.ToInt32(dr["ACTIVITY_ID"].ToString());
                            tr.TreasureName = dr["ACTIVITY_NAME"].ToString();
                            returnValue.Add(tr); 
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    public IList GetAllGiftVote(int actID)
    {
        IList returnValue = new ArrayList();
        string sql = "SELECT GIFT.GIFT_ID,GIFT_NAME,GIFT_IMAGE,CASE WHEN B.VOTES IS NULL  THEN 0 ELSE B.VOTES END \"VOTES\" FROM GIFT ";
        sql += " LEFT JOIN (SELECT dbo.LOTTERY_VOTE.GIFT_ID,SUM(VOTES) AS VOTES FROM  LOTTERY_VOTE GROUP BY LOTTERY_VOTE.GIFT_ID)AS B ";
	    sql += "    ON B.GIFT_ID = dbo.GIFT.GIFT_ID ";
        sql += " WHERE  GIFT.ACTIVITY_ID = @ACTIVITY_ID ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actID;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            LotteryGift lotteryGift = new LotteryGift();
                            lotteryGift.GiftId = dr["GIFT_ID"].ToString();
                            lotteryGift.GiftName = dr["GIFT_NAME"].ToString();
                            lotteryGift.Images = dr["GIFT_IMAGE"].ToString();
                            lotteryGift.Votes = Convert.ToInt32(dr["VOTES"].ToString());
                            returnValue.Add(lotteryGift);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    public IList GetUserVotesGift(int actId,int userId)
    {
        IList returnValue = new ArrayList();
        string sql = " SELECT GIFT.GIFT_ID,GIFT_NAME,GIFT_IMAGE,CASE WHEN VOTES IS NULL  THEN 0 ELSE VOTES END \"VOTES\" FROM GIFT ";
        sql += " LEFT JOIN LOTTERY_VOTE ON LOTTERY_VOTE.GIFT_ID = GIFT.GIFT_ID AND LOTTERY_VOTE.ACCOUNT_ID =@ACCOUNT_ID ";
        sql += " WHERE  GIFT.ACTIVITY_ID = @ACTIVITY_ID ";
        string stateString = "";
        if (actId <= 0)
            actId = activity.Id;
        sql = string.Format(sql, stateString);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int).Value = userId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            LotteryGift lotteryGift = new LotteryGift();
                            lotteryGift.GiftId = dr["GIFT_ID"].ToString();
                            lotteryGift.GiftName = dr["GIFT_NAME"].ToString();
                            lotteryGift.Images = dr["GIFT_IMAGE"].ToString();
                            lotteryGift.Votes = Convert.ToInt32(dr["VOTES"].ToString());
                            returnValue.Add(lotteryGift);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    public IList GetGiftVoteUsers(int actId,string giftId,string userName,int pageIndex,int pageSize)
    {
        IList returnValue = new ArrayList();
        string sql = " WITH SUBJECTORDER AS (  ";
        sql += " SELECT LOGIN_ID,REALNAME,ACCOUNT.ACCOUNT_ID,NICKNAME,VOTES,ROW_NUMBER() OVER (ORDER BY VOTES DESC) AS SN ";
        sql += " FROM ACCOUNT LEFT JOIN LOTTERY_VOTE ON ACCOUNT.ACCOUNT_ID = LOTTERY_VOTE.ACCOUNT_ID ";
        sql += " WHERE GIFT_ID = @GIFT_ID AND ACTIVITY_ID = @ACTIVITY_ID {0} )";
        sql += " SELECT * FROM SUBJECTORDER WHERE SN BETWEEN @PAGEINDEX AND @PAGESIZE ;";
        string queryMember = "";
        if (!string.IsNullOrEmpty(userName))
            queryMember = " AND ( ACCOUNT.LOGIN_ID LIKE @MEMBERID OR ACCOUNT.REALNAME LIKE @MEMBERID OR ACCOUNT.NICKNAME LIKE @MEMBERID ) ";
        sql = string.Format(sql, queryMember);
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@GIFT_ID", SqlDbType.VarChar, 50).Value = giftId;
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                cmd.Parameters.Add("@PAGEINDEX", SqlDbType.Int).Value = (1 + (pageIndex - 1) * pageSize);
                cmd.Parameters.Add("@PAGESIZE", SqlDbType.Int).Value = (pageIndex) * pageSize;
                cmd.Parameters.Add("@MEMBERID", SqlDbType.NVarChar, 50).Value = "%"+userName+"%";
                try
                {
                    conn.Open();
                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        while (dr.Read())
                        {
                            Treasure_Top treasureTop = new Treasure_Top();
                            treasureTop.LoginId = dr["LOGIN_ID"].ToString();
                            treasureTop.RealName = dr["REALNAME"].ToString();
                            treasureTop.NickName = dr["NICKNAME"].ToString();
                            treasureTop.AccountId = Convert.ToInt32(dr["ACCOUNT_ID"].ToString());
                            treasureTop.Lottery = Convert.ToInt32(dr["VOTES"].ToString());
                            returnValue.Add(treasureTop);
                        }
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    public int GetVoteCountById(int actId,string giftId)
    {
        int returnValue = 0;
        string sql = "SELECT COUNT(ACCOUNT_ID) AS VOTES  FROM LOTTERY_VOTE WHERE ACTIVITY_ID = @ACTIVITY_ID AND GIFT_ID =@GIFT_ID ";
        using (SqlConnection conn = TreasureHuntDB.createConnection())
        {
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@ACTIVITY_ID", SqlDbType.Int).Value = actId;
                cmd.Parameters.Add("@GIFT_ID", SqlDbType.VarChar).Value = giftId;
                try
                {
                    conn.Open();
                    object o = cmd.ExecuteScalar();
                    if (!(o is DBNull))
                    {
                       returnValue = Convert.ToInt32(o); ;
                    }
                }
                catch (Exception ex)
                {
                    log4.Error(ex.ToString());
                    throw ex;
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
        return returnValue;
    }

    #region 檢查程式
    /// <summary>
    /// 檢查是否一日超過三套
    /// </summary>
    /// <param name="aid"></param>
    /// <returns></returns>
    public DataTable CheckDailyGetSet(int aid)
    {
        string sql = @"
        select ddd.*,ac.login_id,ac.realname,ac.nickname from (
        select activity_id,account_id,treasure_id,states,count(treasure_id) as daily_count,convert(nvarchar,modify_date,111) as modifydate from (
        select activity_id,account_id,treasure_id,states,convert(nvarchar,modify_date,111) as modify_date
         from dbo.TREASURE_HUNT_LOG
         where activity_id = @activity_id and states = 2 group by activity_id,account_id,treasure_id,states,modify_date
        ) as dd   group by activity_id,account_id,treasure_id,states,modify_date
        ) as ddd 
        left join account ac on ac.account_id = ddd.account_id
        where daily_count > @daily_count
        ";
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        var dt = SqlHelper.GetDataTable("TreasureHuntConnString", sql,
           DbProviderFactories.CreateParameter("GSSConnString", "@daily_count", "@daily_count", daily_set),
           DbProviderFactories.CreateParameter("GSSConnString", "@activity_id", "@activity_id", aid));
        
        return dt;
    }
    public DataTable CheckDailyChangeSet(int aid)
    {
        string sql = @"
        select ddd.*,ac.login_id,ac.realname,ac.nickname from (
        select activity_id,account_id,states,sum(daily_count) as daily_count,convert(nvarchar,create_date,111) as create_date from (
        select activity_id,account_id,count(treasure_id) as daily_count,states,convert(nvarchar,create_date,111) as create_date
         from dbo.TREASURE_HUNT_LOG
         where activity_id = @activity_id and states in (3,4) group by activity_id,account_id,states,create_date
        ) as dd   group by activity_id,account_id,states,create_date
        ) as ddd 
        left join account ac on ac.account_id = ddd.account_id
        where daily_count > 2 and states = 3 or daily_count > 1 and states = 4
        order by account_id,create_date,states
        ";
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        var dt = SqlHelper.GetDataTable("TreasureHuntConnString", sql,
           DbProviderFactories.CreateParameter("GSSConnString", "@activity_id", "@activity_id", aid));

        return dt;
    }
    public DataTable CheckUserPackage(int aid)
    {
        string sql = @"
        select * from 
        (
        select ddd.*,ac.login_id,ac.realname,ac.nickname,ug.piece,ddd.daily_count-isnull(eee.daily_count,0)+isnull(fff.daily_count,0) as nowcount,isnull(eee.daily_count,0) as changecount,isnull(fff.daily_count,0) as changegetcount from (
        select activity_id,account_id,states,treasure_id,sum(daily_count) as daily_count from (
        select activity_id,account_id,treasure_id,count(treasure_id) as daily_count,states
         from dbo.TREASURE_HUNT_LOG
         where activity_id = @activity_id and states  =2 group by activity_id,account_id,treasure_id,states
        ) as dd   group by activity_id,account_id,treasure_id,states
        ) as ddd 
        left join (
        select activity_id,account_id,states,treasure_id,sum(daily_count) as daily_count from (
        select activity_id,account_id,treasure_id,count(treasure_id) as daily_count,states
         from dbo.TREASURE_HUNT_LOG
         where activity_id = @activity_id and states  =3 group by activity_id,account_id,treasure_id,states
        ) as dd   group by activity_id,account_id,treasure_id,states
        ) as eee on eee.account_id = ddd.account_id and eee.treasure_id = ddd.treasure_id
        left join
        (
        select activity_id,account_id,states,treasure_id,sum(daily_count) as daily_count from (
        select activity_id,account_id,treasure_id,count(treasure_id) as daily_count,states
         from dbo.TREASURE_HUNT_LOG
         where activity_id = @activity_id and states  =4 group by activity_id,account_id,treasure_id,states
        ) as dd   group by activity_id,account_id,treasure_id,states
        ) as fff  on fff.account_id = ddd.account_id and fff.treasure_id = ddd.treasure_id
        left join dbo.USERS_PACKAGE ug on ug.account_id = ddd.account_id and ug.treasure_id = ddd.treasure_id and ug.activity_id = ddd.activity_id
        left join account ac on ac.account_id = ddd.account_id
        )gggg where piece <> nowcount
        ";
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        var dt = SqlHelper.GetDataTable("TreasureHuntConnString", sql,
           DbProviderFactories.CreateParameter("GSSConnString", "@activity_id", "@activity_id", aid));

        return dt;
    }
    #endregion
    /**
 * 把真實姓名馬賽克
 */
    public string replaceRealName(string nickname, string realname)
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



    class ReadDocumentRecord
    {
        int m_ID;
        int m_Activity_ID;
        int m_Account_ID;
        string m_Url;
        int m_Times;
        string m_Title;
        string m_Unit;
        DateTime m_ModifyDate;
        DateTime m_CreateDate;

        public int ID
        {
            get { return m_ID; }
            set { m_ID = value; }
        }
        public int Activity_ID
        {
            get { return m_Activity_ID; }
            set { m_Activity_ID = value; }
        }
        public int Account_ID
        {
            get { return m_Account_ID; }
            set { m_Account_ID = value; }
        }
        public string Url
        {
            get { return m_Url; }
            set { m_Url = value; }
        }
        public int Times
        {
            get { return m_Times; }
            set { m_Times = value; }
        }
        public string Title
        {
            get { return m_Title; }
            set { m_Title = value; }
        }
        public string Unit
        {
            get { return m_Unit; }
            set { m_Unit = value; }
        }
        public DateTime ModifyDate
        {
            get { return m_ModifyDate; }
            set { m_ModifyDate = value; }
        }
        public DateTime CreateDate
        {
            get { return m_CreateDate; }
            set { m_CreateDate = value; }
        }

    }

    class ProbabilityNumber
    {
        int m_Maximum;
        int m_Minimum;

        public int Maximum
        {
            get { return m_Maximum; }
            set { m_Maximum = value; }
        }
        public int Minimum
        {
            get { return m_Minimum; }
            set { m_Minimum = value; }
        }
    }

    private enum TreasureHuntState
    {
        Get_Box,
        Faild,
        Get_Treasure,
        Chnage_Treasure,
        Get_Treasure_by_Change
    }

    public class Treasure
    {
        int m_TreasureId;
        string m_TreasureName;
        string m_Icon;
        public int TreasureId
        {
            get { return m_TreasureId; }
            set { m_TreasureId = value; }
        }
        public string TreasureName
        {
            get { return m_TreasureName; }
            set { m_TreasureName = value; }
        }
        public string Icon
        {
            get { return m_Icon; }
            set { m_Icon = value; }
        }
    }

    public class Treasure_Top
    {
        int m_AccountId;
        string m_LoginId;
        string m_RealName;
        string m_NickName;
        int m_TotalSuit;
        string m_Qualifications;
        int m_Lottery;
        string m_TotalReadCount;
        int m_Total_Treasure;
        int m_Change_Treasure;
        int m_Real_Treasure;
        public int AccountId
        {
            get { return m_AccountId; }
            set { m_AccountId = value; }
        }
        public string LoginId
        {
            get { return m_LoginId; }
            set { m_LoginId = value; }
        }
        public string RealName
        {
            get { return m_RealName; }
            set { m_RealName = value; }
        }
        public string NickName
        {
            get { return m_NickName; }
            set { m_NickName = value; }
        }
        public int TotalSuit
        {
            get { return m_TotalSuit; }
            set { m_TotalSuit = value; }
        }
        public string Qualifications
        {
            get { return m_Qualifications; }
            set { m_Qualifications = value; }
        }
        public int Lottery
        {
            get { return m_Lottery; }
            set { m_Lottery = value; }
        }
        public string TotalReadCount
        {
            get { return m_TotalReadCount; }
            set { m_TotalReadCount = value; }
        }
        public int Total_Treasure
        {
            get { return m_Total_Treasure; }
            set { m_Total_Treasure = value; }
        }
        public int Change_Treasure
        {
            get { return m_Change_Treasure; }
            set { m_Change_Treasure = value; }
        }
        public int Real_Treasure
        {
            get { return m_Real_Treasure; }
            set { m_Real_Treasure = value; }
        }
    }

    public class TreasureLogDetal
    {
        string m_UnitName;
        string m_DocumentTitle;
        string m_Url;
        string m_TreasureName;
        string m_GetDate;
        int m_States;
        public string UnitName
        {
            get { return m_UnitName; }
            set { m_UnitName = value; }
        }
        public string DocumentTitle
        {
            get { return m_DocumentTitle; }
            set { m_DocumentTitle = value; }
        }
        public string Url
        {
            get { return m_Url; }
            set { m_Url = value; }
        }
        public string TreasureName
        {
            get { return m_TreasureName; }
            set { m_TreasureName = value; }
        }
        public string GetDate
        {
            get { return m_GetDate; }
            set { m_GetDate = value; }
        }
        public int States
        {
            get { return m_States; }
            set { m_States = value; }
        }
    }

    public class LotteryGift
    {
        string m_GiftName;
        string m_GiftId;
        string m_Images;
        int m_Votes;
        public string GiftName
        {
            get { return m_GiftName; }
            set { m_GiftName = value; }
        }
        public string GiftId
        {
            get { return m_GiftId; }
            set { m_GiftId = value; }
        }
        public string Images
        {
            get { return m_Images; }
            set { m_Images = value; }
        }
        public int Votes
        {
            get { return m_Votes; }
            set { m_Votes = value; }
        }
    }
}
