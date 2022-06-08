using System;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Data.SqlClient;
using System.Data;
using System.Transactions;
using System.Collections.Generic;
using System.Configuration;

[WebService(Namespace = "http://tempuri.org/")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class KnowledgeForumService : System.Web.Services.WebService
{
    //預設文件編輯者
    private string _gMemberId = null;
    //資料庫連線字串
    private string _gConnectionString = null;
    //限定的使用者IP(以';'作為分隔)
    private string _gUserIPWhiteList = ConfigurationManager.AppSettings["userIPWhiteList"].Trim();

    //建構子
    public KnowledgeForumService()
    {
        //檢查是否允許使用這個web service
        if (!IsUserAuthorization())
            throw new Exception("Access Service Denied. The request is from " + HttpContext.Current.Request.UserHostAddress);         

        //初始化參數
        _gMemberId = ConfigurationManager.AppSettings["DefaultEditor"].Trim();
        _gConnectionString = ConfigurationManager.ConnectionStrings["CoaConnString"].ConnectionString.Trim();
    }
    //取得資料庫連線
    private SqlConnection GetSQLConnection()
    {
        return new SqlConnection(_gConnectionString);
    }
    //允許授權使用的IP名單
    private List<string> AuthorizationIPs
    {
        get
        {
            string[] whiteIPs = _gUserIPWhiteList.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> resultList = new List<string>(whiteIPs.Length);
            if(whiteIPs.Length>0)
                resultList.AddRange(whiteIPs);
            return resultList;
        }
    }
    //用來判斷使用者是不是可以獲得授權使用這個web service
    private bool IsUserAuthorization()
    {
        bool isUserAuthorization = false;
        //判斷IP是不是處在白名單內
        isUserAuthorization = AuthorizationIPs.Contains(HttpContext.Current.Request.UserHostAddress);
        return isUserAuthorization;
    }
    //驗證字串類型參數
    private void validationStringParameter(string value)
    {
        if (string.IsNullOrEmpty(value.Trim())) 
            throw new Exception("string is not null or empty");
    }
    //驗正整數類型參數
    private void validationIntParameter(int value)
    {
        if (value < 0) 
            throw new Exception("number greater than zero");
    }
    //驗證知識家問題主鍵值
    private void validationQuestionId(int iQuestionId)
    {
        if (!IsExistQuestionId(iQuestionId))
            throw new Exception("question id does not exist");
    }
    //判斷知識家問題主鍵值是否存在
    private bool IsExistQuestionId(int iQuestionId)
    {
        //文件使用者ID
        string memberId = _gMemberId;

        using (SqlConnection connection = GetSQLConnection())
        {
            bool result = false;
            string sqlString = string.Empty;
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            try
            {
                connection.Open();
                sqlString = @"SELECT * FROM CuDTGeneric 
                    WHERE  (iCUItem = @iCUItem) --AND (iEditor = @iEditor)  
                    AND (iBaseDSD = N'39') AND (iCTUnit = N'932') AND (xKeyword LIKE '%安全農業%')";

                command.Parameters.Add("@iCUItem", SqlDbType.Int).Value = iQuestionId;
                command.Parameters.Add("@iEditor", SqlDbType.NVarChar).Value = memberId;
                command.CommandText = sqlString;
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count == 1) 
                    result = true;
                return result;
            }
            catch (SoapException soapEx)
            {
                throw soapEx;
            }
            catch (Exception)
            {
                throw new Exception("Function Exception");
            }
        }
    }

    /// <summary>
    /// 新增知識家問題
    /// </summary>
    /// <param name="sTitle">標題</param>
    /// <param name="sContent">內文</param>
    /// <returns>新增的問題主鍵值</returns>
    [WebMethod]
    public int InsertQuestion(string sTitle, string sContent)
    {
        //驗證參數
        validationStringParameter(sTitle);
        validationStringParameter(sContent);

         //文件使用者ID
        string memberId = _gMemberId;
        int documentId = -1;

        TransactionOptions transOption = new TransactionOptions();
        transOption.IsolationLevel = System.Transactions.IsolationLevel.Serializable;

        using (TransactionScope transScope = new TransactionScope())
        {
            using (SqlConnection connection = GetSQLConnection())
            {
                string sqlString = string.Empty;
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                try
                {
                    connection.Open();

                    //step1.新增文件資料
                    sqlString =
                        @"INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, topCat, xKeyword, xBody, siteId) 
                        VALUES(N'39', N'932', N'Y', @sTitle, @iEditor, 0, N'A', N'安全農業;', @xbody,N'3') 
                        SELECT IDENT_CURRENT('CuDTGeneric') AS documentid";

                    command.Parameters.Add("@sTitle", SqlDbType.NVarChar).Value = sTitle;
                    command.Parameters.Add("@iEditor", SqlDbType.NVarChar).Value = memberId;
                    command.Parameters.Add("@xbody", SqlDbType.NText).Value = sContent;

                    command.CommandText = sqlString;
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(command);
                    da.Fill(dt);
                    documentId = int.Parse(dt.Rows[0]["documentid"].ToString());
                    command.Parameters.Clear();

                    //step2.新增文件擴充資料
                    sqlString =
                        @"INSERT INTO KnowledgeForum(gicuitem) VALUES (@gicuitem)";

                    command.Parameters.Add("@gicuitem", SqlDbType.Int).Value = documentId;
                    command.CommandText = sqlString;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step3.
                    sqlString =
                        @"UPDATE ActivityMemberNew SET QuestionGrade = QuestionGrade + 2 WHERE MemberId = @memberid";

                    command.CommandText = sqlString;
                    command.Parameters.Add("@memberid", SqlDbType.NVarChar).Value = memberId;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step4.
                    sqlString = @"INSERT INTO MemberGradeShare (memberId, shareDate) VALUES(@memberId, GETDATE())";

                    command.CommandText = sqlString;
                    command.Parameters.Add("@memberId", SqlDbType.NVarChar).Value = memberId;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step5.
                    sqlString =
                        @"UPDATE MemberGradeShare SET shareAsk = shareAsk + 2  WHERE memberId = @memberId AND CONVERT(varchar, shareDate, 111) = CONVERT(varchar, GETDATE(), 111)";

                    command.CommandText = sqlString;
                    command.Parameters.Add("@memberId", SqlDbType.NVarChar).Value = memberId;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step6.
                    sqlString =
                        @"INSERT INTO MemberGradeShareDetail VALUES(@memberId, GETDATE(), N'shareAsk', 2, 0, 'Y',@documentId,'0')";

                    command.CommandText = sqlString;
                    command.Parameters.Add("@memberId", SqlDbType.NVarChar).Value = memberId;
                    command.Parameters.Add("@documentId", SqlDbType.Int).Value = documentId;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    transScope.Complete();
                    return documentId;
                }
                catch (SoapException soapEx)
                {
                    throw soapEx;
                }
                catch (Exception)
                {
                    throw new Exception("Function Exception");
                }
            }
        }
    }

    /// <summary>
    /// 更新知識家問題
    /// </summary>
    /// <param name="iQuestionId">問題主鍵值</param>
    /// <param name="sTitle">標題</param>
    /// <param name="sContent">內文</param>
    /// <returns>異動資料筆數</returns>
    [WebMethod]
    public int UpdateQuestion(int iQuestionId, string sTitle, string sContent)
    {
        //驗證參數
        validationIntParameter(iQuestionId);
        validationStringParameter(sTitle);
        validationStringParameter(sContent);
        validationQuestionId(iQuestionId);

        //文件使用者ID
        string memberId = _gMemberId;

        int resultValue = 0;
        TransactionOptions transOption = new TransactionOptions();
        transOption.IsolationLevel = System.Transactions.IsolationLevel.Serializable;

        using (TransactionScope transScope = new TransactionScope())
        {
            using (SqlConnection connection = GetSQLConnection())
            {
                string sqlString = string.Empty;
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                try
                {
                    connection.Open();
                    //step1.更新文件資料
                    sqlString = @"UPDATE CuDTGeneric SET dEditDate=GETDATE(),sTitle = @sTitle,xBody = @xbody 
                        WHERE  (iCUItem = @iCUItem) AND (iEditor = @iEditor)  AND (iBaseDSD = N'39') 
                        AND (iCTUnit = N'932') AND (xKeyword LIKE '%安全農業;%')";

                    command.Parameters.Add("@iCUItem", SqlDbType.Int).Value = iQuestionId;
                    command.Parameters.Add("@iEditor", SqlDbType.NVarChar).Value = memberId;
                    command.Parameters.Add("@sTitle", SqlDbType.NVarChar).Value = sTitle;
                    command.Parameters.Add("@xbody", SqlDbType.NText).Value = sContent;
                    command.CommandText = sqlString;
                    resultValue = command.ExecuteNonQuery();
                    command.Parameters.Clear();
                    transScope.Complete();
                    return resultValue;
                }
                catch (SoapException soapEx)
                {
                    throw soapEx;
                }
                catch (Exception)
                {
                    throw new Exception("Function Exception");
                }
            }
        }
    }

    /// <summary>
    /// 新增知識家討論
    /// </summary>
    /// <param name="iQuestionId">討論所屬的問題主鍵值</param>
    /// <param name="sContent">內文</param>
    /// <returns>新增的討論主鍵值</returns>
    [WebMethod]
    public int InsertDiscussion(int iQuestionId, string sContent)
    {
        //驗證參數
        validationIntParameter(iQuestionId);
        validationStringParameter(sContent);
        validationQuestionId(iQuestionId);

        //文件使用者ID
        string memberId = _gMemberId;

        TransactionOptions transOption = new TransactionOptions();
        transOption.IsolationLevel = System.Transactions.IsolationLevel.Serializable;

        using (TransactionScope transScope = new TransactionScope())
        {
            using (SqlConnection connection = GetSQLConnection())
            {
                string sqlString = string.Empty;
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                try
                {
                    connection.Open();

                    //step1.新增文件資料
                    sqlString =
                        @"INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, topCat, xKeyword, xBody, siteId) 
                        VALUES(N'39', N'933', N'Y', @sTitle, @iEditor, 0, N'A', N'安全農業;', @xbody,N'3') 
                        SELECT IDENT_CURRENT('CuDTGeneric') AS documentid";

                    command.Parameters.Add("@sTitle", SqlDbType.NVarChar).Value = "討論-" + iQuestionId.ToString();
                    command.Parameters.Add("@iEditor", SqlDbType.NVarChar).Value = memberId;
                    command.Parameters.Add("@xbody", SqlDbType.NText).Value = sContent;

                    command.CommandText = sqlString;
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(command);
                    da.Fill(dt);
                    int subDocumentId = int.Parse(dt.Rows[0]["documentid"].ToString());
                    command.Parameters.Clear();

                    //step2.新增文件擴充資料
                    sqlString =
                        @"INSERT INTO KnowledgeForum(gicuitem, ParentIcuitem) VALUES (@gicuitem, @parenticuitem)";

                    command.Parameters.Add("@gicuitem", SqlDbType.Int).Value = subDocumentId;
                    command.Parameters.Add("@parenticuitem", SqlDbType.Int).Value = iQuestionId;
                    command.CommandText = sqlString;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step3.更新統計次數
                    sqlString =
                        @"UPDATE KnowledgeForum SET DiscussCount = DiscussCount + 1 WHERE gicuitem = @gicuitem";

                    command.Parameters.Add("@gicuitem", SqlDbType.Int).Value = iQuestionId;
                    command.CommandText = sqlString;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step4.
                    sqlString =
                        @"UPDATE MemberGradeShare SET shareDiscuss = shareDiscuss + 2  
                        WHERE memberId = @memberId AND CONVERT(varchar, shareDate, 111) = CONVERT(varchar, GETDATE(), 111)";

                    command.Parameters.Add("@memberId", SqlDbType.NVarChar).Value = memberId;
                    command.CommandText = sqlString;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step5.
                    sqlString =
                        @"INSERT INTO MemberGradeShareDetail VALUES(@memberId, GETDATE(), 'shareDiscuss', 2, 4, 'Y',@gicuitem,'0')";

                    command.Parameters.Add("@memberId", SqlDbType.NVarChar).Value = memberId;
                    command.Parameters.Add("@gicuitem", SqlDbType.Int).Value = subDocumentId;
                    command.CommandText = sqlString;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    //step6.
                    sqlString =
                        @"UPDATE ActivityMemberNew SET DiscussGrade = DiscussGrade + 2 WHERE MemberId = @memberId";

                    command.Parameters.Add("@memberId", SqlDbType.NVarChar).Value = memberId;
                    command.CommandText = sqlString;
                    command.ExecuteNonQuery();
                    command.Parameters.Clear();

                    transScope.Complete();
                    //回傳討論主鍵值
                    return subDocumentId;
                }
                catch (SoapException soapEx)
                {
                    throw soapEx;
                }
                catch (Exception)
                {
                    throw new Exception("Function Exception");
                }
            }
        }
    }

    /// <summary>
    /// 更新知識家討論
    /// </summary>
    /// <param name="iDiscussionId">討論主鍵值</param>
    /// <param name="sContent">內文</param>
    /// <returns>異動資料筆數</returns>
    [WebMethod]
    public int UpdateDiscussion(int iDiscussionId, string sContent)
    {
        //驗證參數
        validationIntParameter(iDiscussionId);
        validationStringParameter(sContent);

        //文件使用者ID
        string memberId = _gMemberId;

        int resultValue = 0;
        TransactionOptions transOption = new TransactionOptions();
        transOption.IsolationLevel = System.Transactions.IsolationLevel.Serializable;

        using (TransactionScope transScope = new TransactionScope())
        {
            using (SqlConnection connection = GetSQLConnection())
            {
                string sqlString = string.Empty;
                SqlCommand command = new SqlCommand();
                command.Connection = connection;
                try
                {
                    connection.Open();
                    //step1.更新文件資料
                    sqlString = @"UPDATE CuDTGeneric SET dEditDate=GETDATE(),xBody = @xbody 
                        WHERE  (iCUItem = @iCUItem) AND (iEditor = @iEditor)  AND (iBaseDSD = N'39') 
                        AND (iCTUnit = N'933') AND (xKeyword LIKE '%安全農業;%')";
                    command.Parameters.Add("@iCUItem", SqlDbType.Int).Value = iDiscussionId;
                    command.Parameters.Add("@iEditor", SqlDbType.NVarChar).Value = memberId;
                    command.Parameters.Add("@xbody", SqlDbType.NText).Value = sContent;
                    command.CommandText = sqlString;
                    resultValue = command.ExecuteNonQuery();
                    command.Parameters.Clear();
                    transScope.Complete();
                    return resultValue;
                }
                catch (SoapException soapEx)
                {
                    throw soapEx;
                }
                catch (Exception)
                {
                    throw new Exception("Function Exception");
                }
            }
        }
    }

}
