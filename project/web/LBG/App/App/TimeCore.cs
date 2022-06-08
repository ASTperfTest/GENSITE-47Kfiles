using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace TimeCore
{
    public class TimeCore
    {

        private long tick = 1;
        private int oneHoursTimedOut = 0;
        private int oneDaysTimedOut = 0;
        private int fiveDaysTimedOut = 0;

        public delegate void timerEvent(string eventType, int executingStatus);
        public event timerEvent accounce;

        private bool after5Days = false;
        private string _connStr = "";

        public string connStr
        {
            get
            {
                return _connStr;
            }
            set
            {
                _connStr = value;
            }
        }


        private SqlConnection openConn()
        {
            SqlConnection connection = new SqlConnection(strConnString());
            try
            {
                connection.Open();
            }
            catch
            {
                return null;
            }
            if (connection.State == ConnectionState.Open) return connection;
            else return null;

        }

        private string strConnString()
        {
          //  return "Data Source=localhost;Initial Catalog=LITTLE_BEAN;Persist Security Info=True;User ID=sa;Password=itc";
            return _connStr;
        }




        public TimeCore(string conn)
        {
        }

        private void OneHourExecute(int seconds)
        {
            accounce("SP_CLOCK", 1);
            SP_CLOCK(seconds);
            accounce("SP_CLOCK", 0);
        }

        
        private void OneDayExecute(int seconds)
        {
            //先選出符合條件者的SESSION_ID
            //1.未被停權 (不為-1者)
            //2.植物有在生長狀態者 (STATUS = 1)
            //3.時間已經超過seconds者
            //var qualified = from session in DB.LBG_SESSION
            //                join session_n2m in DB.LBG_GUEST_N2M on session.UID equals session_n2m.TOID
            //                join guest in DB.LBG_GUEST on session_n2m.FROMID equals guest.UID
            //                where session.STATUS == 1 && guest.STATUS != -1
            //                select new { session.UID, session.DAYS, session.LAST_CHECK_STATE, session_n2m.FROMID };


            string SQL = "select  session.UID, session.DAYS, session.LAST_CHECK_STATE, session_n2m.FROMID  from LBG_GUEST_N2M session_n2m , LBG_SESSION session ,LBG_GUEST guest ";
            SQL = SQL + " where session_n2m.TOID=session.UID and session_n2m.FROMID = guest.UID";
            SQL = SQL + " and session.STATUS = 1 and guest.STATUS <> -1 ";
            SqlConnection connection = this.openConn();
            if (connection == null) return ;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = SQL;
            cmd.Connection = connection;
            SqlDataReader result = cmd.ExecuteReader();
            if (!result.Read()) return;
            try
            {
                
                do
                {
                    DateTime dt = System.Convert.ToDateTime(result["LAST_CHECK_STATE"].ToString());
                    TimeSpan ts = (DateTime.Now - dt);
                    if (ts.TotalSeconds < seconds) continue;
                    //guest的environment的clock為0者
                    string UID = result["UID"].ToString ();
                    string FROMID = result["FROMID"].ToString();
                    SqlConnection connection2 = this.openConn();

                    string SQL2 = "select o.CLOCK  from LBG_GUEST_N2M n2m , LBG_ENVIRONMENT o ";
                    SQL2 = SQL2 + " where o.UID = n2m.TOID and o.CLOCK = 0 and n2m.FROMID = '" + FROMID + "'";
                    if (connection2 == null) return;
                    SqlCommand cmd2 = new SqlCommand();
                    cmd2.CommandText = SQL2;
                    cmd2.Connection = connection2;
                    SqlDataReader result2 = cmd2.ExecuteReader();
                    if (!result2.Read()) return;

                    try
                    {
                        //第六日起 , 每5n+1日進行一次蟲害評估
                        int DAYS = System.Convert.ToInt32(result["DAYS"]);
                        int CLOCK = System.Convert.ToInt32(result2["CLOCK"]);
                        if (((DAYS % 5) == 1) && (DAYS > 1)) after5Days = true;
                        try
                        {
                            if (CLOCK != 0) continue;
                        }
                        catch
                        {
                            continue;
                        }

                        if (after5Days)
                        {
                            //第六天
                            //執行蟲害鑑定
                            Random rnd = new Random();
                            int rndValue = rnd.Next(1, 2);  //蟲1 or 蟲2;
                            SP_BUG(UID, rndValue); //值蟲程序
                            after5Days = false;
                            continue;
                        }

                    }
                    finally
                    {
                        cmd2.Cancel();
                        result2.Close();
                        connection2.Close();
                    }
                    accounce("SP_WATER_DECAY", 1);
                    SP_WATER_DECAY(UID);
                    accounce("SP_WATER_DECAY", 0);

                    accounce("SP_FERTITY_DECAY", 1);
                    SP_FERTITY_DECAY(UID);
                    accounce("SP_FERTITY_DECAY", 0);

                    accounce("SP_WEATHER", 1);
                    SP_WEATHER(seconds);
                    accounce("SP_WEATHER", 0);

                    accounce("SP_CHECK_DEAD", 1);
                    SP_CHECK_DEAD(seconds);
                    accounce("SP_CHECK_DEAD", 0);

                    accounce("SP_UPDATE_SESSION_TXN_TIMESTAMP", 1);
                    SP_UPDATE_SESSION_TXN_TIMESTAMP(UID);
                    accounce("SP_UPDATE_SESSION_TXN_TIMESTAMP", 0);

                }
                while (result.Read());
            }
            catch
            {
                return;
            }
            finally
            {
                cmd.Cancel();
                result.Close();
                connection.Close();
            }




        }



        public int hours
        {
            get
            {
                return oneHoursTimedOut;
            }
            set
            {
                oneHoursTimedOut = value;
                oneDaysTimedOut = oneHoursTimedOut * 24;
                fiveDaysTimedOut = oneDaysTimedOut * 5;
            }
        }

        //1 Tick為 1遊戲小時
        public void OneTick()
        {
            //執行一遊戲小時的呼叫
            accounce("SP_CLOCK", 1);
            SP_CLOCK(oneHoursTimedOut);
            accounce("SP_CLOCK", 0);

            //因為以下這些與CLOCK有關 , 所以要每遊戲小時呼叫一次
            //SP_SCORE自己會去找出DAYS % 5 ==0者
            accounce("SP_SCORE", 1);
            SP_SCORE(oneHoursTimedOut); //至少間隔一個遊戲小時
            accounce("SP_SCORE", 0);

            //SP_STAR會自己去找出DAYS == 45n or 55n or 75n者
            accounce("SP_STAR", 1);
            SP_STAR();
            accounce("SP_STAR", 0);
            OneDayExecute(oneHoursTimedOut);
        }

        public void SP_BUG(string UID, int bugType)
        {
            string SQL = "SP_BUG";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@SESSION_ID", SqlDbType.NVarChar, 50);
                parameter.Value = UID;

                parameter = cmd.Parameters.Add(
                "@BUGTYPE", SqlDbType.Int );
                parameter.Value = bugType;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();

            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_WATER_DECAY(string UID)
        {
            string SQL = "SP_WATER_DECAY";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@SESSION_ID", SqlDbType.NVarChar, 50);
                parameter.Value = UID;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();

            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_FERTITY_DECAY(string UID)
        {
            string SQL = "SP_FERTITY_DECAY";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@SESSION_ID", SqlDbType.NVarChar, 50);
                parameter.Value = UID;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();

            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_WEATHER(int seconds)
        {
            string SQL = "SP_WEATHER";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@TIMEDOUT", SqlDbType.Int);
                parameter.Value = seconds;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();

            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_STAR()
        {
            string SQL = "SP_STAR";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlDataReader result = cmd.ExecuteReader();
                result.Close();

            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_CHECK_DEAD(int seconds)
        {
            string SQL = "SP_CHECK_DEAD";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@TIMEDOUT", SqlDbType.Int );
                parameter.Value = seconds;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();

            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_UPDATE_SESSION_TXN_TIMESTAMP(string UID)
        {
            string SQL = "SP_UPDATE_SESSION_TXN_TIMESTAMP";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@SESSION_ID", SqlDbType.NVarChar , 50);
                parameter.Value = UID;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();

            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_CLOCK(int seconds)
        {
            string SQL = "SP_CLOCK";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {
                
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@TIMEDOUT", SqlDbType.Int);
                parameter.Value = seconds;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();
                
            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }

        public void SP_SCORE(int timedOut)
        {

            //DB_Function.SP_CLOCK(seconds);
            string SQL = "SP_SCORE";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {
                
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
                SqlParameter parameter = cmd.Parameters.Add(
                "@TIMEDOUT", SqlDbType.Int);
                parameter.Value = timedOut;
                SqlDataReader result = cmd.ExecuteReader();
                result.Close();
                
            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }
        }


        public void SP_START()
        {
            string SQL = "SP_STAR";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {
                
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;

                SqlDataReader result = cmd.ExecuteReader();
                result.Close();
                
            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }

        }

        public void ranking()
        {
            
            //DB_Function.SP_SCORE_RANKING();
            string SQL = "SP_SCORE_RANKING";
            SqlConnection connection = this.openConn();
            if (connection == null) return;
            SqlCommand cmd = new SqlCommand();
            try
            {
                accounce("SP_SCORE_RANKING", 1);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = SQL;
                cmd.Connection = connection;
               
                SqlDataReader result = cmd.ExecuteReader();
                result.Close();
                accounce("SP_SCORE_RANKING", 0);
            }
            catch
            {
                return;
            }
            finally
            {
                connection.Close();
            }

        }
    }
}
