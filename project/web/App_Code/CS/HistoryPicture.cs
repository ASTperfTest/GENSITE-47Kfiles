using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using log4net;
using GSS.Vitals.COA.Data;
using System.Data.SqlClient;
using System.Data;
using System.Collections;

/// <summary>
/// HistoryPicture 的摘要描述
/// </summary>
public class HistoryPicture
{
    public string loginId = "";
    public Activity activity;
    public string userRealname ="";
    public string userNickname = "";
    public string userEmail = "";
    private int accountId;
    private ILog log4 = LogManager.GetLogger("DR");
    public int ctRootId = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["ctRootId"]);
    public int currentNodeId = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["currentNodeId"]);
    public int iCTUnitPic = Convert.ToInt32(System.Web.Configuration.WebConfigurationManager.AppSettings["iCTUnitPic"]);
    public HistoryPicture(string login_id)
    {
        log4net.Config.XmlConfigurator.Configure();
        GSS.Vitals.COA.Data.DbConnectionHelper.LoadSetting();
        if (!string.IsNullOrEmpty(login_id) )
        {
            loginId = login_id.Trim();
            accountId = account_id;
        }
        activity = getActivity();
    }
    #region 活動日期
    public void SetActivity(int actId)
    {
        activity = new Activity();
        activity.Id = actId;
    }

    public class Activity
    {
        int m_Id;
        int m_Activity_Type;
        public int Id
        {
            get { return m_Id; }
            set { m_Id = value; }
        }
        public int Activity_Type
        {
            get { return m_Activity_Type; }
            set { m_Activity_Type = value; }
        }

    }
    
    public int getactivity_id()
    {
        return activity.Id;
    }

    private Activity getActivity()
    {

        Activity act = new Activity();
        string sql = "SELECT ACTIVITY_ID,ACTIVITY_TYPE FROM ACTIVITY WHERE GETDATE() BETWEEN ONLINE_DATE AND OFFLINE_DATE";
        try
        {
            var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql);
            if (dt.Rows.Count > 0)
            {
                var dr = dt.Rows[0];
                act.Id = Convert.ToInt32(dr["ACTIVITY_ID"]);
                act.Activity_Type = Convert.ToInt32(dr["ACTIVITY_TYPE"]);
            }
            else
            {
                act = null;
            }
        }
        catch (Exception ex)
        {
            log4.Error(ex.ToString());
            throw ex;
        }
        
        return act;

    }

    
    #endregion

    #region 帳號
    /**
     * 取得account_id
     */
    public int account_id
    {
        get
        {
            int id = 0;

            string sql = "SELECT ACCOUNT_ID FROM ACCOUNT WHERE login_id = @LOGIN_ID";
            try
            {
                var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
                DbProviderFactories.CreateParameter("HistoryPictureConnString", "@LOGIN_ID", "@LOGIN_ID", loginId));
                if (dt.Rows.Count > 0)
                {
                    id = Convert.ToInt32(dt.Rows[0]["ACCOUNT_ID"].ToString());
                }
            }
            catch (Exception ex)
            {
                log4.Error(ex.ToString());
                throw ex;
            }

            return id;
        }
    }

    public bool CanNotPlayGame(string login_id)
    {
        string sql = "select disable from ACCOUNT where login_id = @login_id ";
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", login_id));
        if (dt.Rows.Count >0)
        {
            return bool.Parse(dt.Rows[0]["disable"].ToString());
        }
        return false;
    }
    public void loginUpdate(string meid)
    {
        string sql = "select account,realname,nickname,email from member where account = @account";
        var dt = SqlHelper.GetDataTable("GSSConnString", sql,
             DbProviderFactories.CreateParameter("GSSConnString", "@account", "@account", meid));
        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            if (dr["nickname"] != null)
            {
                userNickname = dr["nickname"].ToString();
            }
            userEmail = dr["email"].ToString();
            userRealname = dr["realname"].ToString();
            loginUpdate(userNickname, dr["realname"].ToString(), dr["email"].ToString());
        }

    }

    /**
     * 登入時更新account
     */
    public void loginUpdate(string nickname, string realname, string email)
    {
        string sql = "SELECT ACCOUNT_ID,REALNAME,NICKNAME,EMAIL FROM ACCOUNT WHERE LOGIN_ID = @LOGIN_ID";
        try
        {
            var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@LOGIN_ID", "@LOGIN_ID", loginId));
            int account_id = 0;
            if ((dt.Rows.Count > 0))
            {
                var dr = dt.Rows[0];
                // 有資料，檢查資料是否變更並更新
                account_id = Convert.ToInt32(dr["ACCOUNT_ID"]);
                string oldNickname = Convert.ToString(dr["NICKNAME"]);
                string oldRealname = Convert.ToString(dr["REALNAME"]);
                string oldEmail = Convert.ToString(dr["EMAIL"]);
                if (nickname.CompareTo(oldNickname) != 0 || realname.CompareTo(oldRealname) != 0 || email.CompareTo(oldEmail) != 0)
                {
                    using (SqlCommand cmd_login = new SqlCommand())
                    {
                        sql = "UPDATE ACCOUNT SET REALNAME=@REAL_NAME, NICKNAME=@NICK_NAME,EMAIL=@EMAIL, MODIFY_DATE=@MODIFY_DATE WHERE ACCOUNT_ID=@ACCOUNT_ID";
                        SqlHelper.ExecuteNonQuery("HistoryPictureConnString", sql,
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@REAL_NAME", "@REAL_NAME", realname),
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@NICK_NAME", "@NICK_NAME", nickname),
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@EMAIL", "@EMAIL", email),
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@MODIFY_DATE", "@MODIFY_DATE", DateTime.Now),
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@ACCOUNT_ID", "@ACCOUNT_ID", account_id));
                    }
                }
            }
            else
            {
                // 無資料，寫入
                sql = "INSERT INTO ACCOUNT(LOGIN_ID, CREATE_DATE, REALNAME, NICKNAME, EMAIL,disable) ";
                sql += " VALUES(@LOGIN_ID, GETDATE() , @REAL_NAME, @NICK_NAME, @EMAIL,0) ";
                sql += " Select @@Identity  ";
                
                object id =SqlHelper.ReturnScalar("HistoryPictureConnString", sql,
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@REAL_NAME", "@REAL_NAME", realname),
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@NICK_NAME", "@NICK_NAME", nickname),
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@EMAIL", "@EMAIL", email),
                            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@LOGIN_ID", "@LOGIN_ID", loginId));
                account_id = int.Parse(id.ToString());
            }
            InsertOtherData(account_id);
        }
        catch (Exception ex)
        {
            log4.Error(ex.ToString());
            throw ex;
        }
    }

    private void InsertOtherData(int id)
    {
        string sql = @"
        if not exists(
        select * from USER_QUESTION_NOW where ACCOUNT_ID = @ACCOUNT_ID and ACTIVITY_ID = 1
        )
        begin 
        Insert into USER_QUESTION_NOW(ACCOUNT_ID,ACTIVITY_ID,Times,nodes,picturemap,STATES,nodes_org,current_node,icuitem)
                    Values(@ACCOUNT_ID,1,0,'',0,@STATES,'',0,0)
        end
        ";
        SqlHelper.ExecuteNonQuery("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@ACCOUNT_ID", "@ACCOUNT_ID", id),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@STATES", "@STATES", (int)HistoryPictureState.New));
    }
    #endregion
    #region 照片

    public string GetUserCurrentPic(int id,int refid)
    {
        string returnValue = "";
        string sql = @"select 
	        * 
        from CuDTGeneric 
        where iCTUnit = @iCTUnit   --要依照實際的單元別而定
        and refId = @refId
        and  iCuitem = @iCuitem";
        var dt = SqlHelper.GetDataTable("GSSConnString",sql,
            DbProviderFactories.CreateParameter("GSSConnString", "@iCTUnit", "@iCTUnit", iCTUnitPic),
             DbProviderFactories.CreateParameter("GSSConnString", "@refId", "@refId", refid),
             DbProviderFactories.CreateParameter("GSSConnString", "@iCuitem", "@iCuitem", id));
        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            returnValue = dr["xImgFile"].ToString();
        }
        return returnValue;
    }

    public string GetCurrentPic(int id)
    {
        string returnValue = "";
        string sql = @"select 
	        * 
        from CuDTGeneric 
        where iCTUnit = @iCTUnit   --要依照實際的單元別而定
        and  iCuitem = @iCuitem";
        var dt = SqlHelper.GetDataTable("GSSConnString", sql,
            DbProviderFactories.CreateParameter("GSSConnString", "@iCTUnit", "@iCTUnit", iCTUnitPic),
             DbProviderFactories.CreateParameter("GSSConnString", "@iCuitem", "@iCuitem", id));
        if (dt.Rows.Count > 0)
        {
            DataRow dr = dt.Rows[0];
            returnValue = dr["xImgFile"].ToString();
        }
        return returnValue;
    }
    public int GetNewQuestions(string nickname, string realname, string email)
    {
        int nodeid = 0;
        string sql = @"
        DECLARE @dailyCount int 
        set @dailyCount = (
        select count(*)  from HISTORY_PICTURE..HISTORY_PICTURE_LOG where (states = 1 or states = 3  ) and ACCOUNT_ID = @ACCOUNT_ID 
           {0} 
        )
        if (@dailyCount < 30) 
        begin
        select 
	        top 1 * 
        from CuDTGeneric GG
        where iCTUnit = @iCTUnit   
        and fCTUPublic = 'Y'
        and refId in (select * from  HISTORY_PICTURE..picture_node)
        and GG.icuitem is not null 
        and GG.icuitem not in (
        select rdr.icuitem from HISTORY_PICTURE..HISTORY_PICTURE_LOG rdr
        left join HISTORY_PICTURE..ACCOUNT as ac on rdr.ACCOUNT_ID = ac.account_id
        where ac.LOGIN_ID = @LOGIN_ID and ((rdr.states = 1 and rdr.CREATE_DATE between ";
        sql += " CONVERT(varchar(12) , DateAdd(\"d\",-1,getdate()) , 111 )  and  CONVERT(varchar(12) , DateAdd(\"d\",1,getdate()) , 111 )) or (states = 3 ) ) )order by newid() ";
        sql += " end ";
        string temps = "and CREATE_DATE between CONVERT(varchar(12) , getdate() , 111 )  and   CONVERT(varchar(12) , DateAdd(\"d\",1,getdate()) , 111 )";
        sql = string.Format(sql, temps);
        try
        {
            var dt = SqlHelper.GetDataTable("GSSConnString", sql,
            DbProviderFactories.CreateParameter("GSSConnString", "@LOGIN_ID", "@LOGIN_ID", loginId),
            DbProviderFactories.CreateParameter("GSSConnString", "@iCTUnit", "@iCTUnit", iCTUnitPic),
            DbProviderFactories.CreateParameter("GSSConnString", "@ACCOUNT_ID", "@ACCOUNT_ID", account_id));
            if ((dt.Rows.Count > 0))
            {
                var dr = dt.Rows[0];
                nodeid = int.Parse(dr["refID"].ToString());
                sql = @"
                    declare @ctRootId int      set @ctRootId=@@ctRootId;
                    declare @currentNodeId int set @currentNodeId=@@currentNodeId;
	
                    with node_Tree(ctNodeId, dataParent, CatName, DataLevel, fullPath) as
                    (
	                    select 
		                        ctNodeId
		                    ,dataParent
		                    ,CatName
		                    ,DataLevel
		                    ,cast('' as varchar(max))
	                    from CatTreeNode
	                    where ctNodeId = @currentNodeId and ctRootID = @ctRootId
	                    union all
	                    select 
		                        CatTreeNode.ctNodeId
		                    ,node_Tree.ctNodeId
		                    ,CatTreeNode.CatName
		                    ,CatTreeNode.DataLevel
		                    ,cast(node_Tree.fullPath + '/' + CatTreeNode.CatName as varchar(max))
	                    from CatTreeNode
	                    inner join node_Tree on CatTreeNode.dataParent = node_Tree.ctNodeId
                    )
                    select top 29 * from node_Tree 
                    where ctNodeid <> @current_node and ctNodeid in (select * from  HISTORY_PICTURE..picture_node)  order by newid()
                ";
                dt = SqlHelper.GetDataTable("GSSConnString", sql,
                    DbProviderFactories.CreateParameter("GSSConnString", "@current_node", "@current_node", nodeid),
                    DbProviderFactories.CreateParameter("GSSConnString", "@@ctRootId", "@@ctRootId", ctRootId),
                    DbProviderFactories.CreateParameter("GSSConnString", "@@currentNodeId", "@@currentNodeId", currentNodeId));
                string othernode = nodeid.ToString();
                string[] ot = new string[dt.Rows.Count+1];
                ot[0] = nodeid.ToString();
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow drr in dt.Rows)
                    {
                        ot[i] = drr["ctNodeId"].ToString();
                        //othernode += ",";
                        //othernode += drr["ctNodeId"].ToString();
                        i++;
                    }
                }
                
                var loo = from p in ot.ToList() 
                          orderby Guid.NewGuid()
                          select p ;
                foreach (string i in loo)
                {
                    if (!string.IsNullOrEmpty(othernode) && !string.IsNullOrEmpty(i))
                        othernode += ",";
                    if (!string.IsNullOrEmpty(i))
                    othernode += i;
                }
                Random rnd = new Random();
                int kmindex = rnd.Next(0, 5);
                int temp = 0;
                int picturemap = 0;
                foreach (int value in Enum.GetValues(typeof(PictureType)))
                {
                    if (kmindex == temp)
                    {
                        picturemap = value;
                        break;
                    }
                    temp++;
                }
                sql = @"
                    BEGIN TRAN
                    INSERT INTO HISTORY_PICTURE_LOG 
                    (ACTIVITY_ID,ACCOUNT_ID,icuitem,STATES,nodes,MESSAGE,picturemap,MODIFY_DATE,CREATE_DATE)
                         VALUES (@ACTIVITY_ID,@ACCOUNT_ID,@icuitem,@STATES,@nodes,@MESSAGE,@picturemap,getdate(),getdate())
                    update USER_QUESTION_NOW  set Times = 0,STATES = @STATES,picturemap=@picturemap,nodes=@nodes,nodes_org=@nodes,icuitem=@icuitem,current_node=@current_node where ACCOUNT_ID = @ACCOUNT_ID and ACTIVITY_ID = @ACTIVITY_ID 
                    commit
                    ";
                int states = (int)HistoryPictureState.Get_Qqestion;
                SqlHelper.ExecuteNonQuery("HistoryPictureConnString", sql,
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@ACTIVITY_ID", "@ACTIVITY_ID", 1),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@ACCOUNT_ID", "@ACCOUNT_ID", account_id),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@icuitem", "@icuitem", int.Parse(dr["icuitem"].ToString())),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@STATES", "@STATES", states),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@MESSAGE", "@MESSAGE", Enum.GetName(typeof(HistoryPictureState), (int)HistoryPictureState.Get_Qqestion)),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@nodes", "@nodes", othernode),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@picturemap", "@picturemap", picturemap),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@current_node", "@current_node", nodeid));
            }
            else
            {
                return 0;
            }

        }
        catch (Exception ex)
        {
            log4.Error(ex.ToString());

            throw ex;
        }
        return nodeid;
    }
    public UserQuestionNow GetUserQuestionNow()
    {
        string sql = @"
            select * from USER_QUESTION_NOW 
            where account_id = @account_id and activity_id = @activity_id 
        ";
        try
        {
            var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
                DbProviderFactories.CreateParameter("HistoryPictureConnString", "@activity_id", "@activity_id", 1),
                DbProviderFactories.CreateParameter("HistoryPictureConnString", "@account_id", "@account_id", account_id));
            UserQuestionNow returnValue = new UserQuestionNow();
            if (dt.Rows.Count > 0)
            {
                var dr = dt.Rows[0];
                
                returnValue.Times = int.Parse(dr["Times"].ToString());
                returnValue.STATES = int.Parse(dr["STATES"].ToString());
                returnValue.Picturemap = int.Parse(dr["picturemap"].ToString());
                returnValue.NodesOrg = dr["nodes_org"].ToString();
                returnValue.Nodes = dr["nodes"].ToString();
                returnValue.Icuitem = int.Parse(dr["icuitem"].ToString());
                returnValue.CurrentNode = int.Parse(dr["current_node"].ToString());
                return returnValue;
            }
        }
        catch (Exception e)
        {
        }
        return null;
    }


    public bool HitHint(int nodeId)
    {
        UserQuestionNow uqn = GetUserQuestionNow();
        if (uqn != null && uqn.STATES != 3 &&uqn.STATES!= 1)
        {
            //if (nodeId.CompareTo(uqn.CurrentNode) == 0)
            //{
            //    int score = 0;
            //    switch (uqn.Times)
            //    {
            //        case 0:
            //            score = 6;
            //            break;
            //        case 1:
            //            score = 5;
            //            break;
            //        case 2:
            //            score = 4;
            //            break;
            //        case 3:
            //            score = 3;
            //            break;
            //        case 4 : 
            //            score = 2;
            //            break;
            //        case 5:
            //            score = 1;
            //            break;
            //    }
            //    UpdateLog(3, uqn.Picturemap, uqn.Nodes, uqn.Icuitem, uqn.Times, score);
            //    return true;
            //}
            //else 
            if (uqn.Times < 5)
            {
                IList<int> ili = new List<int>();
                foreach (int value in Enum.GetValues(typeof(PictureType)))
                {
                    if ((value & uqn.Picturemap) != value)
                    {
                        ili.Add(value);
                    }
                }
                var dt = from p in ili.ToList()
                         orderby Guid.NewGuid()
                         select p;
                uqn.Picturemap = uqn.Picturemap + dt.First();
                string[] nodess = uqn.Nodes.Split(',');
                int ddeletecount = 4;
                if (nodeId.CompareTo(uqn.CurrentNode) == 0)
                    ddeletecount = 5;
                string[] ds = (from p in nodess.ToList()
                          orderby Guid.NewGuid()
                          where p != uqn.CurrentNode.ToString()
                          select p).Take(5).ToArray();


                string[] dd = (from p in nodess.ToList()
                         where !ds.Contains(p)
                         select p).ToArray();


                string temp = "";
                //if (nodeId.CompareTo(uqn.CurrentNode) == 0)
                //{
                    foreach (string s in dd)
                    {
                        if (!string.IsNullOrEmpty(temp)) temp += ",";
                        temp += s;
                    }
                //}
                //else
                //{
                //    foreach (string s in dd)
                //    {
                //        if (string.Compare(s, nodeId.ToString(), true) != 0)
                //        {
                //            if (!string.IsNullOrEmpty(temp))
                //                temp += ",";
                //            temp += s;
                //        }
                //    }
                //}
                uqn.Nodes = temp;
                uqn.STATES = (int)HistoryPictureState.Hit_Hint;
                uqn.Times = uqn.Times + 1;
                UpdateLog(2, uqn.Picturemap, uqn.Nodes, uqn.Icuitem, uqn.Times,0);
                return true;
            }
            //else if (uqn.Times == 5)
            //{
            //    UpdateLog((int)HistoryPictureState.Faild, uqn.Picturemap, uqn.Nodes, uqn.Icuitem, uqn.Times,0);
            //}
        }
        return false;
    }

    public bool CheckAnswer(int nodeId)
    {
        UserQuestionNow uqn = GetUserQuestionNow();
        if (uqn != null && uqn.STATES != 3 && uqn.STATES != 1)
        {
            if (nodeId.CompareTo(uqn.CurrentNode) == 0)
            {
                int score = 0;
                switch (uqn.Times)
                {
                    case 0:
                        score = 6;
                        break;
                    case 1:
                        score = 5;
                        break;
                    case 2:
                        score = 4;
                        break;
                    case 3:
                        score = 3;
                        break;
                    case 4:
                        score = 2;
                        break;
                    case 5:
                        score = 1;
                        break;
                }
                UpdateLog(3, uqn.Picturemap, uqn.Nodes, uqn.Icuitem, uqn.Times, score);
                return true;
            }
            else
            {
                UpdateLog((int)HistoryPictureState.Faild, uqn.Picturemap, uqn.Nodes, uqn.Icuitem, uqn.Times, 0);
            }
        }
        return false;
    }

    private bool UpdateLog(int states, int picturemap, string othernode,int icuitem,int times,int score)
    {
        string sql = @"
             INSERT INTO HISTORY_PICTURE_LOG 
                    (ACTIVITY_ID,ACCOUNT_ID,icuitem,STATES,SCORE,nodes,MESSAGE,picturemap,MODIFY_DATE,CREATE_DATE)
                         VALUES (@ACTIVITY_ID,@ACCOUNT_ID,@icuitem,@STATES,@SCORE,@nodes,@MESSAGE,@picturemap,getdate(),getdate())
            update USER_QUESTION_NOW  set Times = @Times,picturemap=@picturemap,STATES = @STATES,nodes=@nodes where ACCOUNT_ID = @ACCOUNT_ID and ACTIVITY_ID = @ACTIVITY_ID 
        ";
        
        SqlHelper.ExecuteNonQuery("HistoryPictureConnString", sql,
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@ACTIVITY_ID", "@ACTIVITY_ID", 1),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@ACCOUNT_ID", "@ACCOUNT_ID", account_id),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@icuitem", "@icuitem", icuitem),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@STATES", "@STATES", states),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@MESSAGE", "@MESSAGE", Enum.GetName(typeof(HistoryPictureState), states)),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@nodes", "@nodes", othernode),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@picturemap", "@picturemap", picturemap),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@SCORE", "@SCORE", score),
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@Times", "@Times", times));
        return false;
    }
    #endregion

    #region 文章
    public DataTable GetQuestionById(string nodesOrg)
    {
        IList returnValue = new ArrayList();
        string sql = @"declare @ctRootId int      set @ctRootId=@@ctRootId;
            declare @currentNodeId int set @currentNodeId=@@currentNodeId;
	
            with node_Tree(ctNodeId, dataParent, CatName, DataLevel, fullPath) as
            (
	            select 
		             ctNodeId
		            ,dataParent
		            ,CatName
		            ,DataLevel
		            ,cast('' as varchar(max))
	            from CatTreeNode
	            where ctNodeId = @currentNodeId and ctRootID = @ctRootId
	            union all
	            select 
		             CatTreeNode.ctNodeId
		            ,node_Tree.ctNodeId
		            ,CatTreeNode.CatName
		            ,CatTreeNode.DataLevel
		            ,cast(node_Tree.fullPath + '/' + CatTreeNode.CatName as varchar(max))
	            from CatTreeNode
	            inner join node_Tree on CatTreeNode.dataParent = node_Tree.ctNodeId
            )
            select top 30 * from node_Tree 
            where ctNodeId in  
        ";
        sql += " ( " + nodesOrg+")";
        var dt = SqlHelper.GetDataTable("GSSConnString", sql,
            DbProviderFactories.CreateParameter("GSSConnString", "@@ctRootId", "@@ctRootId", ctRootId),
             DbProviderFactories.CreateParameter("GSSConnString", "@@currentNodeId", "@@currentNodeId", currentNodeId));

        return dt;

    }
    #endregion
    #region 排行榜
    public IList GetTop(string loginid, string email, int pageSize, int pageIndex)
    {
        IList returnvalue = new ArrayList();
        string sql = @"
             with subjectOrder as (
            select ROW_NUMBER() OVER (ORDER BY  score desc,counts desc) AS Serial,bb.*,aa.login_id,aa.realname,aa.nickname,aa.disable from (
            select hp.ACCOUNT_ID,sum(SCORE) as score,count(hp.ACCOUNT_ID) as counts from HISTORY_PICTURE_LOG hp
            left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
            where ( hp.states = 3  or hp.states=1 ) and hp.ACCOUNT_ID not in  ( select ACCOUNT_ID from ACCOUNT where Disable = 1 )
            {0} {1}
            group by hp.ACCOUNT_ID
            ) bb
            left join ACCOUNT aa on aa.account_id = bb.account_id
            )
     select *   from subjectOrder where Serial between @pageindex1 and @pageindex2     order by Serial,score desc,counts desc
        ";
        string logquery = "";
        if (loginid != "")
        {
            logquery = "and (a.login_id like @login_id or a.REALNAME like @login_id and a.NICKNAME like @login_id)";
        }
        string emailquery = "";
        if (email != "")
        {
            emailquery = " and ( a.email like @email  ) ";
        }
        sql =string.Format(sql, logquery, emailquery);
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", "'%"+loginid+"%'"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@email", "@email", "'%" + email + "%'"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex1", "@pageindex1", (1 + (pageIndex - 1) * pageSize)),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex2", "@pageindex2", (pageIndex) * pageSize));
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                HistoryTop historyTop = new HistoryTop();
                historyTop.AccountId = int.Parse(dr["ACCOUNT_ID"].ToString());
                historyTop.RealName = dr["realname"].ToString();
                historyTop.NickName = dr["nickname"].ToString();
                historyTop.Count = int.Parse(dr["counts"].ToString());
                historyTop.Score = int.Parse(dr["score"].ToString());
                historyTop.LoginId = dr["login_id"].ToString();
                if (dr["disable"] != null && dr["disable"].ToString() != "")
                {
                    historyTop.disable = bool.Parse(dr["disable"].ToString());
                }
                else
                {
                    historyTop.disable = false;
                }
                returnvalue.Add(historyTop);
            }
        }
        return returnvalue;
    }

    public IList GetTopUnsys(string loginid, string email, int scorelow, int scoreup, int answercountlow, int answercountupp, string startTime, string endTime, int pageSize, int pageIndex)
    {
        IList returnvalue = new ArrayList();
        string sql = @"
            with subjectOrder as (
            select ROW_NUMBER() OVER (ORDER BY score desc,counts desc) AS Serial,bb.*,aa.login_id,aa.realname,aa.nickname,aa.disable from (
            select hp.ACCOUNT_ID,sum(SCORE) as score,count(hp.ACCOUNT_ID) as counts from HISTORY_PICTURE_LOG hp
            left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
            where ( hp.states = 3  or hp.states=1 )
            {0} {1} {4}
            group by hp.ACCOUNT_ID
            ) bb
            left join ACCOUNT aa on aa.account_id = bb.account_id
            where 1=1  {2} {3}
            )
         select *   from subjectOrder where Serial between @pageindex1 and @pageindex2     order by Serial,score desc,counts desc
        ";
        string logquery = "";
        if (loginid != "")
        {
            logquery = "and (a.login_id like @login_id or a.REALNAME like @login_id and a.NICKNAME like @login_id)";
        }
        string emailquery = "";
        if (email != "")
        {
            emailquery = " and ( a.email like @email  ) ";
        }
        string datequery = "";
        if (startTime != "" && endTime != "")
        {
            datequery = " and ( hp.CREATE_DATE between @startTime and @endTime ) ";
        }
        string scorequery = "";
        if (scorelow > 0 || scoreup > 0)
        {
            scorequery = " and ( score between @scorelow and  @scoreup ) ";
        }
        string countsquery = "";
        if (answercountlow > 0 || answercountupp > 0)
        {
            countsquery = " and ( counts between @answercountlow and  @answercountupp ) ";
        }
        sql = string.Format(sql, logquery, emailquery, scorequery, countsquery, datequery);
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", "%" + loginid + "%"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@email", "@email", "%" + email + "%"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scorelow", "@scorelow", scorelow),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scoreup", "@scoreup", scoreup),
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountlow", "@answercountlow", answercountlow),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountupp", "@answercountupp", answercountupp),
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@startTime", "@startTime", startTime),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@endTime", "@endTime", endTime),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex1", "@pageindex1", (1 + (pageIndex - 1) * pageSize)),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex2", "@pageindex2", (pageIndex) * pageSize));
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                HistoryTop historyTop = new HistoryTop();
                historyTop.AccountId = int.Parse(dr["ACCOUNT_ID"].ToString());
                historyTop.RealName = dr["realname"].ToString();
                historyTop.NickName = dr["nickname"].ToString();
                historyTop.Count = int.Parse(dr["counts"].ToString());
                historyTop.Score = int.Parse(dr["score"].ToString());
                historyTop.LoginId = dr["login_id"].ToString();
                if (dr["disable"] != null && dr["disable"].ToString() != "")
                {
                    historyTop.disable = bool.Parse(dr["disable"].ToString());
                }
                else
                {
                    historyTop.disable = false;
                }
                returnvalue.Add(historyTop);
            }
        }
        return returnvalue;
    }

    public int GetTopCoun(string loginid, string email)
    {
        int returnvalue = 0;
        string sql = @"
         select count(*) as count from ( 
        select hp.ACCOUNT_ID,sum(SCORE) as score,count(hp.ACCOUNT_ID) as counts 
			        from HISTORY_PICTURE_LOG hp
                    left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
                    where ( hp.states = 3  or hp.states=1 ) and hp.ACCOUNT_ID not in  ( select ACCOUNT_ID from ACCOUNT where Disable = 1 )
                  {0} {1}
                    group by hp.ACCOUNT_ID
        ) bb
        ";
        string logquery = "";
        if (loginid != "")
        {
            logquery = "and (a.login_id like @login_id' or a.REALNAME like @login_id and a.NICKNAME like @login_id)";
        }
        string emailquery = "";
        if (email != "")
        {
            emailquery = " and ( a.email like @email  ) ";
        }
        sql = string.Format(sql, logquery, emailquery);
        var obj = SqlHelper.ReturnScalar("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", "'%" + loginid + "%'"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@email", "@email", "'%" + email + "%'"));
        if (obj != null)
            returnvalue = int.Parse(obj.ToString());
        return returnvalue;
    }

    public int GetTopCounUnsys(string loginid, string email, int scorelow, int scoreup, int answercountlow, int answercountupp, string startTime, string endTime)
    {
        int returnvalue = 0;
        string sql = @"
            select count(*) from (
            select hp.ACCOUNT_ID,sum(SCORE) as score,count(hp.ACCOUNT_ID) as counts from HISTORY_PICTURE_LOG hp
            left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
            where ( hp.states = 3 or hp.states = 1 )
            {0} {1} {4}
            group by hp.ACCOUNT_ID
            ) bb
            left join ACCOUNT aa on aa.account_id = bb.account_id
            where 1=1  {2} {3}
        ";
        string logquery = "";
        if (loginid != "")
        {
            logquery = "and (a.login_id like @login_id or a.REALNAME like @login_id and a.NICKNAME like @login_id)";
        }
        string emailquery = "";
        if (email != "")
        {
            emailquery = " and ( a.email like @email  ) ";
        }
        string datequery = "";
        if (startTime != "" && endTime != "")
        {
            datequery = " and ( hp.CREATE_DATE between @startTime and @endTime ) ";
        }
        string scorequery = "";
        if (scorelow > 0 || scoreup > 0)
        {
            scorequery = " and ( score between @scorelow and  @scoreup ) ";
        }
        string countsquery = "";
        if (answercountlow > 0 || answercountupp > 0)
        {
            countsquery = " and ( counts between @answercountlow and  @answercountupp ) ";
        }
        sql = string.Format(sql, logquery, emailquery, scorequery, countsquery, datequery);
        var obj = SqlHelper.ReturnScalar("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", "%" + loginid + "%"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@email", "@email", "%" + email + "%"),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scorelow", "@scorelow", scorelow),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@scoreup", "@scoreup", scoreup),
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountlow", "@answercountlow", answercountlow),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@answercountupp", "@answercountupp", answercountupp),
             DbProviderFactories.CreateParameter("HistoryPictureConnString", "@startTime", "@startTime", startTime),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@endTime", "@endTime", endTime));
        if (obj != null)
            returnvalue = int.Parse(obj.ToString());
        return returnvalue;
    }

    public HistoryTop GetUserScoreData(string loginid)
    {
        HistoryTop returnValue = new HistoryTop();
        string sql = @"
            select * from (
            select hp.ACCOUNT_ID,sum(SCORE) as score,count(hp.ACCOUNT_ID) as counts from HISTORY_PICTURE_LOG hp
            left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
            where hp.states = 3 
            and a.login_id = @login_id
            group by hp.ACCOUNT_ID
            ) bb
            left join ACCOUNT aa on aa.account_id = bb.account_id
        ";
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id",loginid));
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                returnValue.AccountId = int.Parse(dr["ACCOUNT_ID"].ToString());
                returnValue.RealName = dr["realname"].ToString();
                returnValue.NickName = dr["nickname"].ToString();
                returnValue.Count = int.Parse(dr["counts"].ToString());
                returnValue.Score = int.Parse(dr["score"].ToString());
            }
        }
        else
        {
            returnValue = null;
        }
        return returnValue;
    }

//    public IList GetUserScoreDataList(string loginid,int pageSize,int pageIndex)
//    {
//        IList returnvalue = new ArrayList();
//        string sql = @"
//            select * from (
//            select hp.ACCOUNT_ID,sum(SCORE) as score,count(hp.ACCOUNT_ID) as counts from HISTORY_PICTURE_LOG hp
//            left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
//            where hp.states = 3 
//            and a.login_id = @login_id
//            group by hp.ACCOUNT_ID
//            ) bb
//            left join ACCOUNT aa on aa.account_id = bb.account_id 
//        ";
//        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql,
//            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", "'%" + loginid + "%'"));
//        if (dt.Rows.Count > 0)
//        {
//            foreach (DataRow dr in dt.Rows)
//            {
//                HistoryTop historyTop = new HistoryTop();
//                historyTop.AccountId = int.Parse(dr["ACCOUNT_ID"].ToString());
//                historyTop.RealName = dr["realname"].ToString();
//                historyTop.NickName = dr["nickname"].ToString();
//                historyTop.Count = int.Parse(dr["counts"].ToString());
//                historyTop.Score = int.Parse(dr["score"].ToString());
//                returnvalue.Add(historyTop);
//            }
//        }
//        return returnvalue;
//    }

    public IList GerUserQuestioninfo(string loginid, int pageSize, int pageIndex)
    {
        IList returnvalue = new ArrayList();
        string sql = @"
        with subjectOrder as (
         select ROW_NUMBER() OVER (ORDER BY hp.CREATE_DATE) AS Serial,hp.ACCOUNT_ID,
			        a.login_id,hp.icuitem,hp.score,convert(nvarchar,hp.create_date,111) as create_time,
			        ct.CatName as stitle ,hp.states
			        from HISTORY_PICTURE_LOG hp
                    left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
			        left join mGIPcoanew..CuDTGeneric cg on cg.icuitem = hp.icuitem
					left join  mGIPcoanew..CatTreeNode ct on cg.refid  = ct.ctNodeId
                    where (hp.states = 3 or hp.states = 1)
                    and a.login_id = @login_id )
        select * from subjectOrder where Serial between @pageindex1 and @pageindex2
        
        ";
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString",sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id",loginid),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex1", "@pageindex1", (1 + (pageIndex - 1) * pageSize)),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex2", "@pageindex2", (pageIndex) * pageSize));
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                HistoryQuestionInfo historyQuestionInfo = new HistoryQuestionInfo();
                historyQuestionInfo.Icuitem = int.Parse(dr["icuitem"].ToString());
                historyQuestionInfo.Score = int.Parse(dr["score"].ToString());
                historyQuestionInfo.CreateDate = dr["create_time"].ToString();
                historyQuestionInfo.Stitle = dr["stitle"].ToString();
                historyQuestionInfo.RowNumber = int.Parse(dr["Serial"].ToString());
                historyQuestionInfo.State = int.Parse(dr["states"].ToString());
                returnvalue.Add(historyQuestionInfo);
            }
        }
        return returnvalue;
    }

    public int GerUserQuestioninfoCount(string loginid)
    {
        int returnvalue =0;
        string sql = @"
         select count(hp.icuitem) as counts
        from HISTORY_PICTURE_LOG hp
        left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
        where (hp.states = 3 or hp.states = 1)
        and a.login_id = @login_id
        
        ";
        var obj = SqlHelper.ReturnScalar("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", loginid));
        if (obj != null)
            returnvalue = int.Parse(obj.ToString());
        return returnvalue;
    }

    public IList GerUserDailyScoreinfo(string loginid, int pageSize, int pageIndex)
    {
        IList returnvalue = new ArrayList();
        string sql = @"
        with subjectOrder as (
         select  ROW_NUMBER() OVER (ORDER BY bb.create_time desc) AS sn, bb.ACCOUNT_ID,bb.login_id,count(bb.login_id) as counts,sum(bb.score) as dailyscore,bb.create_time from (
                    select hp.ACCOUNT_ID,
			        a.login_id,hp.score,convert(nvarchar,hp.create_date,111) as create_time,
			       hp.states
			        from HISTORY_PICTURE_LOG hp
                    left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
                    where (hp.states = 3 or hp.states = 1)
                    and a.login_id = @login_id
					) bb
					group by create_time,login_id,ACCOUNT_ID )
                    select * from subjectOrder where sn between @pageindex1 and @pageindex2
        
        ";
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString",sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id",loginid),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex1", "@pageindex1", (1 + (pageIndex - 1) * pageSize)),
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@pageindex2", "@pageindex2", (pageIndex) * pageSize));
        if (dt.Rows.Count > 0)
        {
            foreach (DataRow dr in dt.Rows)
            {
                HistoryQuestionInfo historyQuestionInfo = new HistoryQuestionInfo();
                historyQuestionInfo.Score = int.Parse(dr["dailyscore"].ToString());
                historyQuestionInfo.CreateDate = dr["create_time"].ToString();
                historyQuestionInfo.RowNumber = int.Parse(dr["sn"].ToString());
                historyQuestionInfo.Icuitem = int.Parse(dr["counts"].ToString());
                returnvalue.Add(historyQuestionInfo);
            }
        }
        return returnvalue;
    }

    public int GerUserDailyScoreinfoCount(string loginid)
    {
        int returnvalue = 0;
        string sql = @"
         select count(*) ascounts from
        (
         select  ROW_NUMBER() OVER (ORDER BY bb.create_time desc) AS sn, bb.ACCOUNT_ID,bb.login_id,count(bb.login_id) as counts,sum(bb.score) as dailyscore,bb.create_time 
				from (
                    select hp.ACCOUNT_ID,
			        a.login_id,hp.score,convert(nvarchar,hp.create_date,111) as create_time,
			       hp.states
			        from HISTORY_PICTURE_LOG hp
                    left join ACCOUNT a on a.ACCOUNT_ID = hp.ACCOUNT_ID 
                    where (hp.states = 3 or hp.states = 1)
                    and a.login_id = @login_id
					) bb
					group by create_time,login_id,ACCOUNT_ID ) hhh
        
        ";
        var obj = SqlHelper.ReturnScalar("HistoryPictureConnString", sql,
            DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", loginid));
        if (obj != null)
            returnvalue = int.Parse(obj.ToString());
        return returnvalue;
    }

    public int GetUserDailyAnswer(string loginid)
    {
        string sql = @"select count(*) as counts  from HISTORY_PICTURE..HISTORY_PICTURE_LOG hp
                left join account aa on aa.account_id = hp.ACCOUNT_ID
                where (states = 1 or states=3) and aa.login_id = @login_id ";
                sql +=" and hp.CREATE_DATE between CONVERT(varchar(12) , getdate() , 111 )  and   CONVERT(varchar(12) , DateAdd(\"d\",1,getdate()) , 111 )";
                var obj = SqlHelper.ReturnScalar("HistoryPictureConnString", sql,
                    DbProviderFactories.CreateParameter("HistoryPictureConnString", "@login_id", "@login_id", loginid));
                if (obj != null)
                {
                    return int.Parse(obj.ToString());
                }
        return 0;
    }

    public DataTable GetOverGetQuestion(int orders)
    {
        string sql = @" select * from (
                    select account_id,count(account_id) as usercount,createdate from (
                    select account_id ,states ,convert(nvarchar,create_date,111) as createdate from dbo.HISTORY_PICTURE_LOG where (states = 1 or states = 3)
                    ) bb
                    group by createdate,account_id
                    )ee
                    left join account a on a.account_id = ee.account_id
                    where usercount>30 {0}
        ";
        string orderstring = "";
        switch (orders)
        {
            case 0:
                orderstring = " order by create_date ,login_id ";
                break;
            case 1:
                orderstring = " order by login_id ,create_date ";
                break;
            default:
                break;

        }
        sql = string.Format(sql, orderstring);
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString",sql);
        return dt;
    }

    public DataTable GetOverGetScore(int orders)
    {
        string sql = @" select * from (
                    select account_id,count(account_id) as usercount,sum(score) as score,createdate from (
                    select account_id ,states ,score,convert(nvarchar,create_date,111) as createdate from dbo.HISTORY_PICTURE_LOG where states = 3
                    ) bb
                    group by createdate,account_id
                    )ee
                    left join account a on a.account_id = ee.account_id
                    where score > 180   {0}
        ";
        string orderstring = "";
        switch (orders)
        {
            case 0:
                orderstring = " order by create_date ,login_id ";
                break;
            case 1:
                orderstring = " order by login_id ,create_date ";
                break;
            default:
                break;

        }
        sql = string.Format(sql, orderstring);
        var dt = SqlHelper.GetDataTable("HistoryPictureConnString", sql);
        return dt;
    }
    #endregion 
    private enum HistoryPictureState
    {
        Get_Qqestion,
        Faild,
        Hit_Hint,
        Success,
        New
    }

    public class UserQuestionNow
    {
        int m_Times;
        string m_nodes;
        int m_picturemap;
        int m_STATES;
        string m_nodes_org;
        int m_current_node;
        int m_icuitem;
        public int Times
        {
            get { return m_Times; }
            set { m_Times = value; }
        }
        public string Nodes
        {
            get { return m_nodes; }
            set { m_nodes = value; }
        }
        public int Picturemap
        {
            get { return m_picturemap; }
            set { m_picturemap = value; }
        }
        public int STATES
        {
            get { return m_STATES; }
            set { m_STATES = value; }
        }
        public string NodesOrg
        {
            get { return m_nodes_org; }
            set { m_nodes_org = value; }
        }
        public int CurrentNode
        {
            get { return m_current_node; }
            set { m_current_node = value; }
        }
        public int Icuitem
        {
            get { return m_icuitem; }
            set { m_icuitem = value; }
        }
    }

    public class HistoryTop
    {
        int m_AccountId;
        string m_LoginId;
        string m_RealName;
        string m_NickName;
        int m_Score;
        int m_Count;
        bool m_disable;
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
        public int Score
        {
            get { return m_Score; }
            set { m_Score = value; }
        }
        public int Count
        {
            get { return m_Count; }
            set { m_Count = value; }
        }
        public bool disable
        {
            get { return m_disable; }
            set { m_disable = value; }
        }
    }
    public class HistoryQuestionInfo
    {
        int m_RowNumber;
        int m_Icuitem;
        int m_Score;
        int m_State;
        string m_CreateDate;
        string m_Stitle;
        public int RowNumber
        {
            get { return m_RowNumber; }
            set { m_RowNumber = value; }
        }
        public int Icuitem
        {
            get { return m_Icuitem; }
            set { m_Icuitem = value; }
        }
        public int Score
        {
            get { return m_Score; }
            set { m_Score = value; }
        }
        public int State
        {
            get { return m_State; }
            set { m_State = value; }
        }
        public string CreateDate
        {
            get { return m_CreateDate; }
            set { m_CreateDate = value; }
        }
        public string Stitle
        {
            get { return m_Stitle; }
            set { m_Stitle = value; }
        }
    }
    public enum PictureType
    {
        ONE = 1,
        TWO = 2,
        THREE = 4,
        FOUR = 8,
        FIVE = 16,
        SIX = 32
    }
}