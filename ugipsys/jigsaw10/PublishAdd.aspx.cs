using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Transactions;
using System.Data.Linq;
using System.Data.Linq.SqlClient;
using System.Net;
using System.Text;
using Jayrock.Json;
using Jayrock.Json.Conversion;
using System.Collections.Specialized;

public partial class PublishAdd : System.Web.UI.Page
{
    public PaginatedList<ResultView> pl;
    public int id;
    public int pid;
    public string topCat;
    public string CtRootId;
    public string CtNodeName;
    public string sTitle;
    public string Status;
    public string strWebUrl;
	public string xPostDateS = "1920/1/1";
	public string xPostDateE= "9999/12/29";
    private IRepository _coa_repository;
    private IRepository _mGIPcoanew_repository;

    public class ReportQueryResult
    {
        public string REPORT_ID { get; set; }
        public string CATEGORY_NAME { get; set; }
        public string CATEGORY_ID { get; set; }
        public string SUBJECT { get; set; }
        public string PUBLISHER { get; set; }
        public DateTime ONLINE_DATE { get; set; }
        public string ACTOR_DETAIL_NAME { get; set; }
        public string STATUS { get; set; }
    }
    public class ResultView
    {
        public int CtRootID { get; set; }
        public int CtNodeID { get; set; }
        public int categoryId { get; set; }
        public int iCUItem { get; set; }
        public string sTitle { get; set; }
        public string CtRootName { get; set; }
        public string CatName { get; set; }
        public string CtUnitName { get; set; }
        public string UserName { get; set; }
        public string dEditDate { get; set; }
        public string deptName { get; set; }
        public char? siteId { get; set; }
        public int? iCTUnit { get; set; }
        public string vGroup { get; set; }
        public char fCTUPublic { get; set; }
    }

    string apiSite = StrFunc.GetWebSetting("LambdaAPISite", "appSettings");//"http://kwpi-coa-kmintra.gss.com.tw/coa/api/";
    string apiKey = StrFunc.GetWebSetting("ApiKey", "appSettings");//"188a0e0577c54d5284f09821091e7fc7";
    string user = StrFunc.GetWebSetting("User", "appSettings");//"localhost";
    int page;
    int pageSize;

    protected void Page_Load(object sender, EventArgs e)
    {
        // LINQ to SQL (Repository)
        _coa_repository = new Repository(new coaDataContext());
        _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());

        // appSettings (Web.config)
        strWebUrl = StrFunc.GetWebSetting("WebURL", "appSettings").Replace("\\", "/");

        // QueryString
        id = int.Parse(Request.QueryString["id"] ?? "0");
        pid = int.Parse(Request.QueryString["pid"] ?? "0");
        topCat = Request.QueryString["topCat"] ?? "";
        //sam 排除已經在農漁地圖的文章用
        var alreadyjigsawid = (from p in _mGIPcoanew_repository.List<KnowledgeJigsaw>()
                               where p.parentIcuitem == id && p.Status == 'Y'
                               select p.ArticleId).ToArray();
        // FormPost
        CtRootId = Request["CtRootId"] ?? "";
        CtNodeName = Request["CtNodeName"] ?? "";
        sTitle = Request["sTitle"] ?? "";
        Status = Request["Status"] ?? "";
		if(!String.IsNullOrEmpty(Request["xPostDateS"]))
		{
			xPostDateS = Request["xPostDateS"];
		}
		if(!String.IsNullOrEmpty(Request["xPostDateE"]))
		{
			xPostDateE = Request["xPostDateE"];
		}
		if (!IsPostBack)
        {
            page = int.Parse(Request.QueryString["page"] ?? "0");
            pageSize = int.Parse(Request.QueryString["pagesize"] ?? "10");

            var rs = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == pid);
            var rs1 = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == id);

            if (rs != null && rs1 != null)
                titles.Text = string.Format("【內容條例清單--{0} {1}】", rs.sTitle, rs1.sTitle);

            switch (CtRootId)
            {
                case "4":
                    #region //previous version.
                    #region //sql
                    //sql += "SELECT DISTINCT REPORT.REPORT_ID, CATEGORY.CATEGORY_NAME, CATEGORY.CATEGORY_ID, REPORT.SUBJECT, REPORT.PUBLISHER, REPORT.ONLINE_DATE, ACTOR_INFO.ACTOR_DETAIL_NAME";
                    //sql += " FROM REPORT";
                    //sql += " INNER JOIN ACTOR_INFO ON REPORT.CREATE_USER = ACTOR_INFO.ACTOR_INFO_ID";
                    //sql += " INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID";
                    //sql += " INNER JOIN CATEGORY ON CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID AND CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID";
                    //sql += " INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID";
                    //sql += " WHERE (REPORT.STATUS = 'PUB') AND (REPORT.ONLINE_DATE < GETDATE()) AND (CATEGORY.DATA_BASE_ID = 'DB020') AND (RESOURCE_RIGHT.ACTOR_INFO_ID IN('001','002'))";
                    //if (Request("CtNodeName") != "")
                    //{sql += "AND (CATEGORY.CATEGORY_NAME LIKE " & pkDate("%" & Request("CtNodeName") & "%", "") & ") "}
                    //if (Request("sTitle") != "")
                    //{sql += "AND (REPORT.SUBJECT LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "}
                    //if (Request("Status") == "Y")
                    //{sql += "AND (REPORT.STATUS = 'PUB') "}
                    //else
                    //{sql += "AND (REPORT.STATUS <> 'PUB') "}
                    //sql += " ORDER BY REPORT.ONLINE_DATE DESC "
                    #endregion
                    //                    string sql = @"SELECT DISTINCT REPORT.REPORT_ID, REPORT.STATUS, REPORT.SUBJECT, REPORT.PUBLISHER, REPORT.ONLINE_DATE, CATEGORY.CATEGORY_NAME, CATEGORY.CATEGORY_ID, ACTOR_INFO.ACTOR_DETAIL_NAME
                    //                                                       FROM REPORT
                    //                                                         INNER JOIN ACTOR_INFO ON REPORT.CREATE_USER = ACTOR_INFO.ACTOR_INFO_ID
                    //                                                         INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID
                    //                                                         INNER JOIN CATEGORY ON CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID AND CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID
                    //                                                         INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID
                    //                                                       WHERE (REPORT.ONLINE_DATE < GETDATE()) AND (CATEGORY.DATA_BASE_ID = 'DB020') AND (RESOURCE_RIGHT.ACTOR_INFO_ID IN('001','002'))
                    //                                                       ORDER BY REPORT.ONLINE_DATE DESC ";
                    //                    using (DataContext _coa = new coaDataContext())
                    //                    {
                    //                        var _result = _coa.ExecuteQuery<ReportQueryResult>(sql);
                    //                        if (CtNodeName != "")
                    //                        {
                    //                            _result = _result.Where(p => p.CATEGORY_NAME.Contains(CtNodeName));
                    //                        }
                    //                        if (sTitle != "")
                    //                        {
                    //                            _result = _result.Where(p => p.SUBJECT.Contains(sTitle));
                    //                        }
                    //                        if (Status == "Y")
                    //                        {
                    //                            _result = _result.Where(p => p.STATUS == "PUB");
                    //                        }
                    //                        else
                    //                        {
                    //                            _result = _result.Where(p => p.STATUS != "PUB");
                    //                        }
                    //                        var result = from p in _result.ToList()
                    //                                     select new ResultView
                    //                                     {
                    //                                         iCUItem = int.Parse(p.REPORT_ID),
                    //                                         CatName = p.CATEGORY_NAME,
                    //                                         //CtNodeID = p.CATEGORY_ID,
                    //                                         sTitle = p.SUBJECT,
                    //                                         deptName = p.PUBLISHER,
                    //                                         dEditDate = String.Format("{0:yyyy/M/d HH:mm}", p.ONLINE_DATE),
                    //                                         UserName = p.ACTOR_DETAIL_NAME,
                    //                                     };
                    //                        //分頁
                    //                        pl = new PaginatedList<ResultView>(result.AsQueryable(), page, pageSize, new string[] { "10", "30", "50" });
                    //                    }
                    #endregion
                    try
                    {
                        JsonArray JOData = (JsonArray)JsonConvert.Import(WebCall_AdvancedSearchresult());
                        JsonArray JODocuments = (JsonArray)JOData[0];
                        int TotalCount = int.Parse(JOData[3].ToString());
                        List<ResultView> result = new List<ResultView>();
                        foreach (JsonObject doc in JODocuments)
                        {
                            JsonArray JOCategories = (JsonArray)doc["Categories"];
                            //JsonArray JOAttributes = (JsonArray)JsonConvert.Import(WebCall_GetDocumentAttribute(doc["UniqueKey"].ToString()));
                            //string _l_publish_units = "";
                            //foreach (JsonObject a in JOAttributes)
                            //{
                            //    if (a["DisplayName"].ToString() == "出版單位")
                            //    {
                            //        _l_publish_units = a["Value"].ToString();
                            //        break;
                            //    }
                            //}
                            if (!alreadyjigsawid.Contains(Convert.ToInt32(doc["UniqueKey"].ToString())))
                            result.Add(
                                new ResultView
                                {
                                    iCUItem = Convert.ToInt32(doc["UniqueKey"].ToString()),
                                    categoryId = int.Parse(JOCategories[0].ToString()),
                                    CatName = "",//WebCall_GetCategoryDisplayNameById(JOCategories[0].ToString()),
                                    //CtNodeID = p.CATEGORY_ID,
                                    sTitle = doc["Title"].ToString(),
                                    deptName = "",//_l_publish_units,
                                    dEditDate = JsonDateTimeUtility.ToDateTime(doc["CreationDatetime"].ToString()).ToShortDateString(),
                                    UserName = WebCall_GetUserDisplayNameBySubjectId(doc["Author"].ToString()),
                                });
                        }
                        //分頁
                        pl = new PaginatedList<ResultView>(result.AsQueryable(), page, pageSize, new string[] { "10", "30", "50" }, result.Count);
                    }
                    catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); Response.End(); }
                    break;
                case "3":
                    #region //sql
                    //sql = ""
                    //sql = sql & " SELECT DISTINCT CatTreeNode.CtRootID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, CatTreeNode.CatName, "
                    //sql = sql & " CtUnit.CtUnitName, CuDTGeneric.dEditDate, Member.realname FROM CuDTGeneric INNER JOIN CatTreeNode ON "
                    //sql = sql & " CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON CatTreeNode.CtRootID = CatTreeRoot.CtRootID "
                    //sql = sql & " INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
                    //sql = sql & " INNER JOIN KnowledgeForum ON CuDTGeneric.icuitem = KnowledgeForum.gicuitem "
                    //sql = sql & " WHERE (CuDTGeneric.siteId = N'3') AND (CatTreeRoot.inUse = 'Y') AND (CuDTGeneric.fCTUPublic = 'Y') "
                    //sql = sql & " AND (CtUnit.inUse = 'Y') AND (CatTreeNode.inUse = 'Y') AND (CatTreeRoot.inUse = 'Y') AND (CatTreeNode.CtUnitID = 932) "
                    //sql = sql & " AND (KnowledgeForum.Status <> 'D')"
                    //If Request("sTitle") <> "" Then
                    //    sql = sql & "AND (CuDTGeneric.sTitle LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "
                    //End If
                    //If Request("Status") = "Y" Then
                    //    sql = sql & "AND (CuDTGeneric.fCTUPublic = 'Y') "
                    //Else
                    //    sql = sql & "AND (CuDTGeneric.fCTUPublic <> 'Y') "
                    //End If
                    //sql = sql & " ORDER BY CuDtGeneric.dEditDate DESC "
                    #endregion
                    {
                        var result = from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                                      join s0 in _mGIPcoanew_repository.List<CatTreeNode>() on p.iCTUnit equals s0.CtUnitID
                                      join s1 in _mGIPcoanew_repository.List<CatTreeRoot>() on s0.CtRootID equals s1.CtRootID
                                      join s2 in _mGIPcoanew_repository.List<CtUnit>() on s0.CtUnitID equals s2.CtUnitID
                                      join s3 in _mGIPcoanew_repository.List<Member>() on p.iEditor equals s3.account
                                      join s4 in _mGIPcoanew_repository.List<KnowledgeForum>() on p.iCUItem equals s4.gicuitem
                                     where (p.showType != '2' || (p.showType == '2' && (SqlMethods.Like(p.xURL,"http://kmweb.coa.gov.tw%") || !SqlMethods.Like(p.xURL,"http://%")))) && !alreadyjigsawid.Contains(p.iCUItem) && (p.siteId == '3') && (s1.inUse == 'Y') && (s2.inUse == 'Y') && (s0.inUse == 'Y') && (s0.CtUnitID == 932) && (s4.Status != 'D') && (Convert.ToDateTime(p.dEditDate) >= Convert.ToDateTime(xPostDateS)) && (Convert.ToDateTime(p.dEditDate) <= (Convert.ToDateTime(xPostDateE)).AddDays(1).AddMinutes(-1))
                                      orderby p.dEditDate descending
                                      select new ResultView
                                      {
                                          CtRootID = s0.CtRootID,
                                          CtNodeID = s0.CtNodeID,
                                          iCUItem = p.iCUItem,
                                          sTitle = p.sTitle,
                                          CtRootName = s1.CtRootName,
                                          CatName = s0.CatName,
                                          CtUnitName = s2.CtUnitName,
                                          UserName = s3.realname,
                                          dEditDate = String.Format("{0:yyyy/M/d HH:mm}", p.dEditDate),
                                          deptName = "知識家",
                                          siteId = p.siteId,
                                          iCTUnit = p.iCTUnit,
                                          vGroup = s1.vGroup,
                                          fCTUPublic = p.fCTUPublic,
                                      };

                        if (sTitle != "")
                        {
                            result = result.Where(p => p.sTitle.Contains(sTitle));
                        }
                        if (Status == "Y")
                        {
                            result = result.Where(p => p.fCTUPublic == 'Y');
                        }
                        else
                        {
                            result = result.Where(p => p.fCTUPublic != 'Y');
                        }

                        //分頁
                        pl = new PaginatedList<ResultView>(result, page, pageSize, new string[] { "10", "30", "50" });
                    }
                    break;
                default:
                    #region //sql
                    //sql = ""	
                    //sql = sql & "SELECT DISTINCT CatTreeNode.CtRootID, CatTreeNode.CtNodeID, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeRoot.CtRootName, "
                    //sql = sql & "CatTreeNode.CatName, CtUnit.CtUnitName, InfoUser.UserName, CuDTGeneric.dEditDate, Dept.deptName FROM CuDTGeneric INNER JOIN "
                    //sql = sql & "CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID INNER JOIN CatTreeRoot ON "
                    //sql = sql & "CatTreeNode.CtRootID = CatTreeRoot.CtRootID INNER JOIN CtUnit ON "
                    //sql = sql & "CatTreeNode.CtUnitID = CtUnit.CtUnitID INNER JOIN InfoUser ON "
                    //sql = sql & "CuDTGeneric.iEditor = InfoUser.UserID INNER JOIN Dept ON CuDTGeneric.iDept = Dept.deptID "
                    //sql = sql & "WHERE (CatTreeRoot.inUse = 'Y') AND (CtUnit.inUse = 'Y') "
                    //sql = sql & "AND (CatTreeRoot.inUse = 'Y') AND (Dept.inUse = 'Y')"
                    //If Request("CtRootId") = "1" Then			
                    //    sql = sql & "AND (CuDtGeneric.siteId = " & pkStr(Request("CtRootId"), "") & ") AND (CatTreeNode.CtRootID = 34) "			
                    //ElseIf Request("CtRootId") = "2" Then			
                    //    sql = sql & "AND (CuDtGeneric.siteId = " & pkStr(Request("CtRootId"), "") & ") AND (CatTreeRoot.vGroup = 'XX')  "								
                    //ElseIf Request("CtRootId") <> "" Then
                    //    sql = sql & "AND (CuDtGeneric.iCtUnit = " & Request("CtRootId") & ") "								
                    //End If
                    //If Request("CtNodeName") <> "" Then
                    //    sql = sql & "AND (CatTreeNode.CatName LIKE " & pkDate("%" & Request("CtNodeName") & "%", "") & ") "
                    //End If	
                    //If Request("sTitle") <> "" Then
                    //    sql = sql & "AND (CuDTGeneric.sTitle LIKE " & pkDate("%" & Request("sTitle") & "%", "") & ") "
                    //End If
                    //If Request("Status") = "Y" Then
                    //    sql = sql & "AND (CuDTGeneric.fCTUPublic = 'Y') "
                    //Else
                    //    sql = sql & "AND (CuDTGeneric.fCTUPublic <> 'Y') "
                    //End If
                    //sql = sql & " ORDER BY CuDtGeneric.dEditDate DESC "
                    #endregion
                    {
                        var result = from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                                     join s0 in _mGIPcoanew_repository.List<CatTreeNode>() on p.iCTUnit equals s0.CtUnitID
                                     join s1 in _mGIPcoanew_repository.List<CatTreeRoot>() on s0.CtRootID equals s1.CtRootID
                                     join s2 in _mGIPcoanew_repository.List<CtUnit>() on s0.CtUnitID equals s2.CtUnitID
                                     join s3 in _mGIPcoanew_repository.List<InfoUser>() on p.iEditor equals s3.UserID
                                     join s4 in _mGIPcoanew_repository.List<Dept>() on p.iDept equals s4.deptID
                                     where (p.showType != '2' || (p.showType == '2' && (SqlMethods.Like(p.xURL,"http://kmweb.coa.gov.tw%") || !SqlMethods.Like(p.xURL,"http://%")))) && !alreadyjigsawid.Contains(p.iCUItem) && (s1.inUse == 'Y') && (s2.inUse == 'Y') && (s4.inUse == 'Y') && (Convert.ToDateTime(p.dEditDate) >= Convert.ToDateTime(xPostDateS)) && (Convert.ToDateTime(p.dEditDate) <= (Convert.ToDateTime(xPostDateE)).AddDays(1).AddMinutes(-1))
                                     orderby p.dEditDate descending
                                     select new ResultView
                                     {
                                         CtRootID = s0.CtRootID,
                                         CtNodeID = s0.CtNodeID,
                                         iCUItem = p.iCUItem,
                                         sTitle = p.sTitle,
                                         CtRootName = s1.CtRootName,
                                         CatName = s0.CatName,
                                         CtUnitName = s2.CtUnitName,
                                         UserName = s3.UserName,
                                         dEditDate = String.Format("{0:yyyy/M/d HH:mm}", p.dEditDate),
                                         deptName = s4.deptName,
                                         siteId = p.siteId,
                                         iCTUnit = p.iCTUnit,
                                         vGroup = s1.vGroup,
                                         fCTUPublic = p.fCTUPublic,
                                     };

                        if (CtRootId == "1")
                        {
                            result = result.Where(p=> p.siteId == CtRootId[0] && p.CtRootID == 34);
                        }
                        else if (CtRootId == "2")
                        {
                            result = result.Where(p => p.siteId == CtRootId[0] && p.vGroup == "XX");
                        }
                        else if (CtRootId != "")
                        {
                            result = result.Where(p => p.iCTUnit == int.Parse(CtRootId));
                        }
                        if (CtNodeName != "")
                        {
                            result = result.Where(p => p.CatName.Contains(CtNodeName));
                        }
                        if (sTitle != "")
                        {
                            result = result.Where(p => p.sTitle.Contains(sTitle));
                        }
                        if (Status == "Y")
                        {
                            result = result.Where(p => p.fCTUPublic == 'Y');
                        }
                        else
                        {
                            result = result.Where(p => p.fCTUPublic != 'Y');
                        }

                        //分頁
                        pl = new PaginatedList<ResultView>(result, page, pageSize, new string[] { "10", "30", "50" });
                    }
                    break;
            }
            //Data Repeater DataBinding
            this.rptList.DataSource = pl;
            this.rptList.DataBind();
        }
        else
        {
            string ckbox = Request.Form["ckbox"] ?? "";
            if (ckbox != "")
            {
                foreach (string x in ckbox.Split(','))
                {
                    var checkExist = _mGIPcoanew_repository.Get<KnowledgeJigsaw>(p => p.Status == 'Y' && p.ArticleId == int.Parse(x) && p.parentIcuitem == id);
                    if (checkExist == null)
                    {
                        string _sTitle = "";
                        string _CtUnitId = "";
                        string _path = "";

                        if (CtRootId == "4")
                        {
                            #region //previous version.
                            #region //sql
                            //sql = sql & " AND (REPORT.REPORT_ID = '" & insert1 & "')"
                            //CtUnitId= rs("CATEGORY_ID")
                            //rs("SUBJECT")

                            //sql2 = "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[dEditDate],[iDept],[showType],[siteId]) "
                            //sql2 = sql2 & "VALUES(" & iBaseDSD & ", " & iCTUnit & ", '" & rs("SUBJECT") & "', '" & IEditor & "', GETDATE(), '" & IDept & "', "
                            //sql2 = sql2 & "'" & showType & "', '" & siteId & "') "
                            //sql2 = "set nocount on;" & sql2 & "; select @@IDENTITY as NewID"	
                            //gicuitem = rs2(0)
                            //sql6 = "INSERT INTO CuDTx7 ([giCuItem]) VALUES(" & gicuitem & ") "		   			

                            //path = "/CatTree/CatTreeContent.aspx?ReportId={0}&DatabaseId=DB020&CategoryId={1}&ActorType=002"
                            //path = replace(path, "{0}", insert1)
                            //path = replace(path, "{1}", CtUnitId)

                            //sql2i = "INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[CtRootId],[CtUnitId],[parentIcuitem],[ArticleId],[Status],[orderArticle],[path]) "
                            //sql2i = sql2i & "VALUES(" & gicuitem & ", " & request("CtRootId") & ", '" & CtUnitId & "', " & Regicuitem & ", " & insert1 & ", 'Y', " & orderArticle & ", '" & path & "')"						
                            #endregion
//                            string sql = @"SELECT DISTINCT REPORT.REPORT_ID, REPORT.STATUS, REPORT.SUBJECT, REPORT.PUBLISHER, REPORT.ONLINE_DATE, CATEGORY.CATEGORY_NAME, CATEGORY.CATEGORY_ID, ACTOR_INFO.ACTOR_DETAIL_NAME
//                                   FROM REPORT
//                                     INNER JOIN ACTOR_INFO ON REPORT.CREATE_USER = ACTOR_INFO.ACTOR_INFO_ID
//                                     INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID
//                                     INNER JOIN CATEGORY ON CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID AND CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID
//                                     INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID
//                                   WHERE (REPORT.ONLINE_DATE < GETDATE()) AND (CATEGORY.DATA_BASE_ID = 'DB020') AND (RESOURCE_RIGHT.ACTOR_INFO_ID IN('001','002'))
//                                   ";
//                            sql += " AND (REPORT.REPORT_ID = '" + x + "') ";
//                            using (DataContext _coa = new coaDataContext())
//                            {
//                                var _result = _coa.ExecuteQuery<ReportQueryResult>(sql);
//                                var result = _result.FirstOrDefault();
//                                _sTitle = result == null ? "" : result.SUBJECT;
//                                _CtUnitId = result == null ? "" : result.CATEGORY_ID;
//                                _path = string.Format("/CatTree/CatTreeContent.aspx?ReportId={0}&DatabaseId=DB020&CategoryId={1}&ActorType=002", x, _CtUnitId);
//                            }
                            #endregion
                            try
                            {
                                JsonObject JOResult = (JsonObject)JsonConvert.Import(WebCall_GetDocumentById(x));
                                JsonArray JOCategories = (JsonArray)JOResult["Categories"];
                                JsonObject JOCategoryFirst = (JsonObject)JOCategories[0];
                                _sTitle = JOResult["VersionTitle"].ToString();
                                _CtUnitId = JOCategoryFirst["CategoryId"].ToString();
                                _path = string.Format("/category/categorycontent.aspx?ReportId={0}&DatabaseId=DB020&CategoryId={1}&ActorType=000", x, _CtUnitId);
                            }
                            catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); Response.End(); }
                        }
                        else
                        {
                            #region //sql
                            //sql1 = "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[dEditDate],[iDept],[showType],[siteId]) "
                            //sql1 = "VALUES(" & iBaseDSD & ", " & iCTUnit & ", '" & rs("sTitle") & "', '" & IEditor & "', GETDATE(), '" & IDept & "', " & "'" & showType & "', '" & siteId & "') "
                            //sql1 = "set nocount on;" & sql1 & "; select @@IDENTITY as NewID"	
                            //gicuitem = rs1(0)
                            //sql6 = "INSERT INTO CuDTx7 ([giCuItem]) VALUES(" & gicuitem & ") "
                            //path = ""			
                            //newsiteid = GetSiteId( insert1 )
                            //newpath = GetPath( newsiteid, insert1 ,rs("CtRootId") ,rs("CtNodeId"))
                            //'response.write gicuitem
                            //sql1i = "INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[CtRootId],[CtUnitId],[parentIcuitem],[ArticleId],[Status],[orderArticle],[path]) "
                            //sql1i = "VALUES(" & gicuitem & ", " & request("CtRootId") & ", '" & request("CtRootId") & "', " & Regicuitem & ", " & insert1 & ", 'Y', " & orderArticle & ", '" & path & "') "
                            #endregion
                            var result = (from p in _mGIPcoanew_repository.List<CuDTGeneric>().Where(p => p.iCUItem == int.Parse(x))
                                         join s in _mGIPcoanew_repository.List<CatTreeNode>() on p.iCTUnit equals s.CtUnitID
                                         select new 
                                         {
                                             p.sTitle,
                                             p.siteId,
                                             p.topCat,
                                             p.xURL,
                                             p.fileDownLoad,
                                             p.showType,
                                             s.CtRootID,
                                             s.CtNodeID,
                                         }).FirstOrDefault();

                            _sTitle = result == null ? "" : result.sTitle;
                            _CtUnitId = CtRootId;
                            _path = getPath(result.siteId.ToString(), int.Parse(x), result.CtRootID, result.CtNodeID, result.topCat, result.showType, result.xURL, result.fileDownLoad);
                        }

                        try
                        {
                            #region //INSERT
                            using (TransactionScope ts = new TransactionScope())
                            {
                                var newCuDTGeneric = new CuDTGeneric
                                {
                                    iBaseDSD = 44,
                                    iCTUnit = 2201,
                                    sTitle = _sTitle,
                                    iEditor = "hyweb",
                                    Created_Date = DateTime.Now,
                                    dEditDate = DateTime.Now,
                                    xPostDate = DateTime.Now,
                                    iDept = "0",
                                    showType = '1',
                                    siteId = '1',
                                };
                                _mGIPcoanew_repository.Create<CuDTGeneric>(newCuDTGeneric);

                                var newCuDTx7 = new CuDTx7
                                {
                                    giCuItem = newCuDTGeneric.iCUItem
                                };
                                _mGIPcoanew_repository.Create<CuDTx7>(newCuDTx7);

                                var newKnowledgeJigsaw = new KnowledgeJigsaw
                                {
                                    gicuitem = newCuDTGeneric.iCUItem,
                                    CtRootId = int.Parse(CtRootId),
                                    CtUnitId = _CtUnitId,
                                    parentIcuitem = id,
                                    ArticleId = int.Parse(x),
                                    Status = 'Y',
                                    orderArticle = 1,
                                    path = _path,
                                };
                                _mGIPcoanew_repository.Create<KnowledgeJigsaw>(newKnowledgeJigsaw);

                                ts.Complete();
                            }
                            #endregion
                        }
                        catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }
                    }
                }
            }
            Redirect();
        }
    }

    private void Redirect()
    {
        StrFunc.GenErrMsg("編修成功！", string.Format("window.location.href=\"PublishQuery.aspx?pid={0}&id={1}&topCat={2}\";", pid, id, topCat));
    }
    public string getCheckbox(int articleid, int id)
    {
        var rs = _mGIPcoanew_repository.Get<KnowledgeJigsaw>(p=> p.Status == 'Y' && p.ArticleId == articleid && p.parentIcuitem == id);
        return string.Format("<input type=\"checkbox\" name=\"ckbox\" value=\"{0}\"{1} />", articleid, rs == null ? "" : " checked=\"checked\"");
    }
    public string getAnchor(string sourceId, int id, int categoryId)
    {
	    string result = "";
        switch (sourceId)
        {
            case "4":
                {
                    result = string.Format("/category/categorycontent.aspx?ReportId={0}&DatabaseId=DB020&CategoryId={1}&ActorType=000", id, categoryId);
                }
                break;
            case "3":
                {
                    var rs = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == id);
                    if (rs != null)
                        result = string.Format("/knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType={1}&CategoryId={2}", id, "A", rs.topCat);
                }
                break;
            case "2":
                {
                    var rs = (from p in _mGIPcoanew_repository.List<CuDTGeneric>().Where(p => p.iCUItem == id)
                              join s0 in _mGIPcoanew_repository.List<CtUnit>() on p.iCTUnit equals s0.CtUnitID
                              join s1 in _mGIPcoanew_repository.List<CatTreeNode>() on s0.CtUnitID equals s1.CtUnitID
                              select new { s1.CtNodeID, s1.CtRootID }).FirstOrDefault();
                    if (rs != null)
                        result = string.Format("/subject/ct.asp?xItem={0}&ctNode={1}&mp={2}", id, rs.CtNodeID, rs.CtRootID);
                }
                break;
            case "1":
                {
                    var rs = (from p in _mGIPcoanew_repository.List<CuDTGeneric>().Where(p => p.iCUItem == id)
                              join s0 in _mGIPcoanew_repository.List<CtUnit>() on p.iCTUnit equals s0.CtUnitID
                              join s1 in _mGIPcoanew_repository.List<CatTreeNode>() on s0.CtUnitID equals s1.CtUnitID
                              select new { s1.CtNodeID }).FirstOrDefault();
                    if (rs != null)
                        result = string.Format("/ct.asp?xItem={0}&ctNode={1}&mp=1", id, rs.CtNodeID);
                }
                break;
            default:
                result = string.Format("/ct.asp?xItem={0}&ctNode={1}&mp=1", id, sourceId);
                break;
        }
        return string.Format("<a href=\"{0}{1}\" target=\"_blank\">View</a>", strWebUrl, result);
    }
    private string getPath(string sourceId, int id, int ctRootId, int CtNodeId, string topCat, char? showType, string xURL, string fileDownLoad)
    {
        string result = "";
        switch (sourceId)
        {
            case "3":
                result = string.Format("/knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType={1}&CategoryId={2}", id, "A", topCat);
                break;
            case "2":
            case "1":
                if (showType == '1')
                {
                    if (sourceId == "2")
                    {
                        var rs = (from p in _mGIPcoanew_repository.List<CuDTGeneric>().Where(p => p.iCUItem == id)
                                  join s0 in _mGIPcoanew_repository.List<CtUnit>() on p.iCTUnit equals s0.CtUnitID
                                  join s1 in _mGIPcoanew_repository.List<CatTreeNode>() on s0.CtUnitID equals s1.CtUnitID
                                  select new { s1.CtNodeID, s1.CtRootID }).FirstOrDefault();
                        if (rs != null)
                            result = string.Format("/subject/ct.asp?xItem={0}&ctNode={1}&mp={2}", id, rs.CtNodeID, rs.CtRootID);
                    }
                    else if (sourceId == "1")
                    {
                        var rs = (from p in _mGIPcoanew_repository.List<CatTreeNode>()
                                  join s0 in _mGIPcoanew_repository.List<CtUnit>() on p.CtUnitID equals s0.CtUnitID
                                  join s1 in _mGIPcoanew_repository.List<CuDTGeneric>() on s0.CtUnitID equals s1.iCTUnit
                                  where p.CtRootID == ctRootId && p.CtNodeID == CtNodeId && s1.iCUItem == id
                                  select new { p.CtNodeID }).FirstOrDefault();
                        if (rs != null)
                            result = string.Format("/ct.asp?xItem={0}&ctNode={1}&mp=1", id, rs.CtNodeID);
                    }
                }
                else if (showType == '2')
                {
                    result = xURL;
                }
                else if (showType == '3')
                {
                    result = string.Format("/public/data/{0}", fileDownLoad);
                }
                break;
        }
        return (result);
    }
    private string WebCall_AdvancedSearchresult()
    {
        string serviceUrl = String.Format(apiSite + "search/advancedresult" + "?") + GetParameterString();
        string result = null;
        
        using (WebClient client = new WebClient())
        {
            try
            {
                client.Encoding = Encoding.UTF8;

                NameValueCollection nc = new NameValueCollection();
                nc["keyword"] = sTitle;
                nc["keywordfield"] = sTitle == "" ? "" : "title";
                nc["folder"] = "";
                nc["category"] = "2";
                nc["docclass"] = "";
                nc["author"] = "";
                nc["datetime"] = "";
                nc["tag"] = "";
                nc["docclassvalue"] = string.Format("公佈日期:[20001231160000 TO {0}] {1} AND (可閱讀分眾導覽\\(前端入口網\\):\"A\" OR 可閱讀分眾導覽\\(前端入口網\\):\"B\" OR 可閱讀分眾導覽\\(前端入口網\\):\"C\")", DateTime.UtcNow.ToString("yyyyMMddHHmmss"), " AND  _l_last_modified_datetime:[" + (Convert.ToDateTime(xPostDateS)).AddHours(-8).ToString("yyyyMMddHHmmss") + " TO " + (Convert.ToDateTime(xPostDateE)).AddDays(1).AddMinutes(-1).AddHours(-8).ToString("yyyyMMddHHmmss") + "]  ");
                nc["containchildfolder"] = "false";
                nc["containchildcategory"] = "true";
                nc["enablekeywordsynonyms"] = "false";
                nc["enabletagsynonyms"] = "false";
				nc["sort"] = "_l_last_modified_datetime";
				
                byte[] bResult = client.UploadValues(serviceUrl, nc);
                string resultData = Encoding.UTF8.GetString(bResult);
                JsonObject JOResult = (JsonObject)JsonConvert.Import(resultData);
                string statuscode = JOResult["statuscode"].ToString();
                if (statuscode == "000")
                {
                    string data = JsonConvert.ExportToString(JOResult["data"]);
                    if (data != null)
                        result = data;
                }
                return result;
            }
            catch (WebException ex)
            {
                throw new Exception(ex.ToString());
            }
        } 
    }
    private string WebCall_GetCategoryDisplayNameById(string id)
    {
        string serviceUrl = String.Format(apiSite + "category/{0}?load_path={1}" + "&", id, "false") + GetParameterString();
        string result = null;

        using (WebClient client = new WebClient())
        {
            try
            {
                client.Encoding = Encoding.UTF8;
                string resultData = client.DownloadString(serviceUrl);
                JsonObject JOResult = (JsonObject)JsonConvert.Import(resultData);
                string statuscode = JOResult["statuscode"].ToString();
                if (statuscode == "000")
                {
                    string data = JsonConvert.ExportToString(JOResult["data"]);
                    JsonObject JOData = (JsonObject)JsonConvert.Import(data);
                    result = JOData["DisplayName"].ToString();
                }
                return result;
            }
            catch
            {
                return "";
            }
        }
    }
    private string WebCall_GetUserDisplayNameBySubjectId(string id)
    {
        string serviceUrl = String.Format(apiSite + "user/exact/{0}" + "?", id) + GetParameterString();
        string result = null;

        using (WebClient client = new WebClient())
        {
            try
            {
                client.Encoding = Encoding.UTF8;
                string resultData = client.DownloadString(serviceUrl);
                JsonObject JOResult = (JsonObject)JsonConvert.Import(resultData);
                string statuscode = JOResult["statuscode"].ToString();
                if (statuscode == "000")
                {
                    string data = JsonConvert.ExportToString(JOResult["data"]);
                    JsonObject JOData = (JsonObject)JsonConvert.Import(data);
                    result = JOData["DisplayName"].ToString();
                }
                return result;
            }
            catch
            {
                return "";
            }
        } 
    }
    private string WebCall_GetDocumentAttribute(string id)
    {
        string serviceUrl = String.Format(apiSite + "document/attribute/{0}?version_number={1}" + "&", id, 1) + GetParameterString();
        string result = null;

        using (WebClient client = new WebClient())
        {
            try
            {
                client.Encoding = Encoding.UTF8;
                string resultData = client.DownloadString(serviceUrl);
                JsonObject JOResult = (JsonObject)JsonConvert.Import(resultData);
                string statuscode = JOResult["statuscode"].ToString();
                if (statuscode == "000")
                {
                    result = JsonConvert.ExportToString(JOResult["data"]);
                }
                return result;
            }
            catch
            {
                return "";
            }
        }
    }
    private string WebCall_GetDocumentById(string id)
    {
        string serviceUrl = String.Format(apiSite + "document/{0}?version_number={1}" + "&", id, 1) + GetParameterString();
        string result = null;

        using (WebClient client = new WebClient())
        {
            try
            {
                client.Encoding = Encoding.UTF8;
                string resultData = client.DownloadString(serviceUrl);
                JsonObject JOResult = (JsonObject)JsonConvert.Import(resultData);
                string statuscode = JOResult["statuscode"].ToString();
                if (statuscode == "000")
                {
                    result = JsonConvert.ExportToString(JOResult["data"]);
                }
                return result;
            }
            catch
            {
                return "";
            }
        } 
    }
    private string GetParameterString()
    {
        return "format=json"
            + "&tid=0"
            + "&who=" + user
            + "&pi=" + page
            + "&ps=" + pageSize
            + "&api_key=" + apiKey;
    }
}
