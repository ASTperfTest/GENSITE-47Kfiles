using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Transactions;
using System.Net;
using System.Text;
using Jayrock.Json;
using Jayrock.Json.Conversion;

public partial class PublishEdit : System.Web.UI.Page
{
    public PaginatedList<ResultView> pl;
    public int id;
    public int pid;
    public string topCat;
    public int page;
    public int pageSize;
    private IRepository _coa_repository;
    private IRepository _mGIPcoanew_repository;
    public class ResultView
    {
        public string sTitle { get; set; }
        public int gicuitem { get; set; }
        public int? orderArticle { get; set; }
        public string xBody { get; set; }
        public string iEditor { get; set; }
        public string xpostdate { get; set; }
        public int? CtRootId { get; set; }
        public string CtUnitId { get; set; }
        public int? ArticleId { get; set; }
        public char? Status { get; set; }
        public int parentIcuitem { get; set; }
        public string dEditDate { get; set; }
        public string path { get; set; }
    }
    string apiSite = StrFunc.GetWebSetting("LambdaAPISite", "appSettings");//"http://kwpi-coa-kmintra.gss.com.tw/coa/api/";
    string apiKey = StrFunc.GetWebSetting("ApiKey", "appSettings");//"188a0e0577c54d5284f09821091e7fc7";
    string user = StrFunc.GetWebSetting("User", "appSettings");//"localhost";

    protected void Page_Load(object sender, EventArgs e)
    {
        // LINQ to SQL (Repository)
        _mGIPcoanew_repository = new Repository(new mGIPcoanewDataContext());
        _coa_repository = new Repository(new coaDataContext());

        // appSettings (Web.config)
        string strWebUrl = StrFunc.GetWebSetting("WebURL", "appSettings").Replace("\\", "/");

        // QueryString
        id = int.Parse(Request.QueryString["id"] ?? "0");
        pid = int.Parse(Request.QueryString["pid"] ?? "0");
        topCat = Request.QueryString["topCat"] ?? "";
        page = int.Parse(Request.QueryString["page"] ?? "0");
        pageSize = int.Parse(Request.QueryString["pagesize"] ?? "10");

        if (!IsPostBack)
        {
            var rs = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == pid);
            var rs1 = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == id);

            if (rs != null && rs1 != null)
                titles.Text = string.Format("【內容條例清單--{0} {1}】", rs.sTitle, rs1.sTitle);

            #region //sql
            //sql2 = "SELECT CuDTGeneric.sTitle, KnowledgeJigsaw.gicuitem as gicuitem, KnowledgeJigsaw.orderArticle, "
            //sql2 += "CuDTGeneric.xBody, CuDTGeneric.iEditor, convert(varchar, CuDTGeneric.xpostdate, 111) as xpostdate, "
            //sql2 += "KnowledgeJigsaw.CtRootId, KnowledgeJigsaw.CtUnitId as CtUnitId, KnowledgeJigsaw.ArticleId as ArticleId, "
            //sql2 += "KnowledgeJigsaw.Status, KnowledgeJigsaw.parentIcuitem, CuDTGeneric.dEditDate, KnowledgeJigsaw.path "
            //sql2 += "FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem "
            //sql2 += "WHERE (KnowledgeJigsaw.Status = 'y') and (KnowledgeJigsaw.parentIcuitem = '"&gicuitem&"') "
            //sql2 += "ORDER BY KnowledgeJigsaw.orderArticle DESC,dEditDate,xpostdate ASC "
            #endregion
            var rs2 = from p in _mGIPcoanew_repository.List<KnowledgeJigsaw>().Where(p => p.parentIcuitem == id && p.Status == 'Y')
                      join s in _mGIPcoanew_repository.List<CuDTGeneric>() on p.gicuitem equals s.iCUItem
                      orderby p.orderArticle descending, s.dEditDate, s.xPostDate ascending
                      select new ResultView
                      {
                          sTitle = s.sTitle,
                          gicuitem = p.gicuitem,
                          orderArticle = p.orderArticle,
                          xBody = s.xBody,
                          iEditor = s.iEditor,
                          xpostdate = String.Format("{0:yyyy/M/d HH:mm}", s.xPostDate),
                          CtRootId = p.CtRootId,
                          CtUnitId = p.CtUnitId,
                          ArticleId = p.ArticleId,
                          Status = p.Status,
                          parentIcuitem = p.parentIcuitem,
                          dEditDate = String.Format("{0:yyyy/M/d HH:mm}", s.dEditDate),
                          path = ((topCat != "E" && topCat != "F") ? strWebUrl : "") + p.path,
                      };

            //分頁
            pl = new PaginatedList<ResultView>(rs2, page, pageSize, new string[] { "10", "30", "50" });

            //Data Repeater DataBinding
            this.rptList.DataSource = pl;
            this.rptList.DataBind();
        }
        else
        {
            switch (Request.Form["action"] ?? "")
            {
                case "addURL": //新增資源推薦聯結

                    #region //sql
                    //orderArticle = "1"
                    //IEditor = "hyweb"
                    //IDept = "0"
                    //showType = "1"
                    //siteId = "1"
                    //iBaseDSD = "44"
                    //iCTUnit = "2201"
                    //CtRootId = 0
                    //insert1=0                             '新增的超連結，沒有原始文章
                    //parentIcuitem=request("gicuitem")     '資源推薦的超連結
                    //
                    //sql2 = "declare @newIDENTITY bigint"
                    //sql2 = sql2 & vbcrlf & "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[dEditDate],[iDept],[showType],[siteId]) "
                    //sql2 = sql2 & vbcrlf & "VALUES(" & iBaseDSD & ", " & iCTUnit & ", '" & title & "', '" & IEditor & "', GETDATE(), '" & IDept & "', '" & showType & "', '" & siteId & "') "
                    //sql2 = sql2 & vbcrlf & "set @newIDENTITY = @@IDENTITY "
                    //sql2 = sql2 & vbcrlf & ""
                    //sql2 = sql2 & vbcrlf & "INSERT INTO CuDTx7 ([giCuItem]) VALUES(@newIDENTITY)"
                    //sql2 = sql2 & vbcrlf & ""
                    //sql2 = sql2 & vbcrlf & "INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[CtRootId],[CtUnitId],[parentIcuitem],[ArticleId],[Status],[orderArticle],[path]) "
                    //sql2 = sql2 & vbcrlf & "VALUES(@newIDENTITY, " & CtRootId & ", 1, " & parentIcuitem & ", " & insert1 & ", 'Y', " & orderArticle & ", '" & Url & "')"						
                    #endregion
                    string Resource_Title = Request.Form["Resource_Title"] ?? "";
                    string Resource_Url = Request.Form["Resource_Url"] ?? "";
                    if (!Resource_Url.StartsWith("http://", true, null))
                        Resource_Url = "http://" + Resource_Url;
                    try
                    {
                        #region //INSERT
                        using (TransactionScope ts = new TransactionScope())
                        {
                            var newCuDTGeneric = new CuDTGeneric
                            {
                                iBaseDSD = 44,
                                iCTUnit = 2201,
                                sTitle = Resource_Title,
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
                                CtRootId = 0,
                                CtUnitId = "1",
                                parentIcuitem = id,
                                ArticleId = 0,
                                Status = 'Y',
                                orderArticle = 1,
                                path = Resource_Url,
                            };
                            _mGIPcoanew_repository.Create<KnowledgeJigsaw>(newKnowledgeJigsaw);

                            ts.Complete();
                        }
                        #endregion
                    }
                    catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }

                    Redirect("編修成功！");
                    break;

                case "del": //刪除選擇

                    try
                    {
                        #region //UPDATE
                        using (TransactionScope ts = new TransactionScope())
                        {
                            string ckbox = Request.Form["ckbox"] ?? "";
                            foreach (string x in ckbox.Split(','))
                            {
                                #region //sql
                                //sql = "UPDATE [mGIPcoanew].[dbo].[KnowledgeJigsaw] SET [Status]='N' WHERE [gicuitem] ='"&update1(i)&"'"
                                //conn.Execute(sql)
                                //'added by Joey,2009/10/12,http://gssjira.gss.com.tw/browse/COAKM-9 , 刪除的同時移除該User的KPI相關數值
                                //'modified by Joey, 2009/10/26, http://gssjira.gss.com.tw/browse/COAKM-19, 當管理員刪除網友留言後，要將[shareJigsaw] 欄位值-1
                                //sql2="SELECT iEditor, convert(varchar, dEditDate, 111) as dEditDate FROM CuDTGeneric WHERE iCUItem=" & update1(i)
                                //set RS = conn.execute(sql2)
                                //if RS.EOF=false then
                                //    sql3="update MemberGradeShare set MemberGradeShare.shareJigsaw=MemberGradeShare.shareJigsaw-1 where MemberGradeShare.memberId='" & RS("iEditor") & "' and convert(varchar, MemberGradeShare.shareDate, 111)='" & RS("dEditDate") & "'"
                                //    conn.Execute(sql3)
                                //end if
                                #endregion
                                var edit_KnowledgeJigsaw = _mGIPcoanew_repository.Get<KnowledgeJigsaw>(p => p.gicuitem == int.Parse(x));
                                if (edit_KnowledgeJigsaw != null)
                                {
                                    edit_KnowledgeJigsaw.Status = 'N';
                                    //'added by Joey, 2009/10/12, http://gssjira.gss.com.tw/browse/COAKM-9 , 刪除的同時移除該User的KPI相關數值
	                                //'modified by Joey, 2009/10/26, http://gssjira.gss.com.tw/browse/COAKM-19, 當管理員刪除網友留言後，要將[shareJigsaw] 欄位值-1
                                    var find_CuDTGeneric = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == int.Parse(x));
                                    if (find_CuDTGeneric != null)
                                    {
                                        var edit_MemberGradeShare = _mGIPcoanew_repository.Get<MemberGradeShare>(p => p.memberId == find_CuDTGeneric.iEditor && p.shareDate == find_CuDTGeneric.dEditDate);
                                        if (edit_MemberGradeShare != null)
                                            edit_MemberGradeShare.shareJigsaw--;
                                    }
                                    _mGIPcoanew_repository.Save();
                                }

                            }
                            ts.Complete();
                        }
                        #endregion
                    }
                    catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }

                    Redirect("刪除成功！");
                    break;

                case "edit":    //編修存檔
                default:
                    
                    #region //sql
                    //sql2="SELECT KnowledgeJigsaw.gicuitem FROM CuDTGeneric INNER JOIN KnowledgeJigsaw ON CuDTGeneric.iCUItem = KnowledgeJigsaw.gicuitem WHERE (KnowledgeJigsaw.Status = 'y') and ( KnowledgeJigsaw.parentIcuitem='"&request("gicuitem")&"')"
                    //set rs2=conn.Execute(sql2)
                    //while not rs2.eof
                    //  if request(rs2("gicuitem"))<>"" then
                    //    sql4="SELECT [orderArticle] FROM [mGIPcoanew].[dbo].[KnowledgeJigsaw] where gicuitem='"&rs2("gicuitem")&"'"
                    //    set rs4=conn.Execute(sql4)
                    //    if isnumeric(request(rs2("gicuitem"))) then        
                    //      if int(rs4(0))<> int(request(rs2("gicuitem"))) then
                    //        sql ="UPDATE [mGIPcoanew].[dbo].[KnowledgeJigsaw] SET [orderArticle] ="&request(rs2("gicuitem"))&"  WHERE gicuitem='"&rs2("gicuitem")&"'"
                    //        conn.Execute(sql)
                    //        sql3="UPDATE [mGIPcoanew].[dbo].[CuDTGeneric] SET [dEditDate]=getdate() WHERE iCUItem ='"&rs2("gicuitem")&"'"
                    //        conn.Execute(sql3)
                    //      end if 
                    //    end if
                    //  end if 
                    //  rs2.movenext
                    //wend
                    #endregion
                    var result = from p in _mGIPcoanew_repository.List<KnowledgeJigsaw>().Where(p => p.Status=='Y' && p.parentIcuitem == id)
                                 join s in _mGIPcoanew_repository.List<CuDTGeneric>() on p.gicuitem equals s.iCUItem
                                 select p.gicuitem;
                    try
                    {
                        #region //UPDATE
                        using (TransactionScope ts = new TransactionScope())
                        {
                            foreach (var x in result)
                            {
                                if (Request.Form[x.ToString()] != null && StrFunc.IsNumeric(Request.Form[x.ToString()]))
                                {
                                    var modified_orderArticle = int.Parse(Request.Form[x.ToString()]);
                                    var original = _mGIPcoanew_repository.Get<KnowledgeJigsaw>(p => p.gicuitem == x);
                                    if (original != null)
                                    {
                                        if (original.orderArticle != modified_orderArticle)
                                        {
                                            original.orderArticle = modified_orderArticle;
                                            var o = _mGIPcoanew_repository.Get<CuDTGeneric>(p => p.iCUItem == x);
                                            if (o != null)
                                            {
                                                o.dEditDate = DateTime.Now;
                                            }
                                            _mGIPcoanew_repository.Save();
                                        }
                                    }
                                }
                            }
                            ts.Complete();
                        }
                        #endregion
                    }
                    catch (Exception ex) { Response.Write(string.Format("<font size='2' color='red'><strong>發生錯誤：{0}</strong></font>", ex)); }

                    Redirect("編修成功！");
                    break;
            }
        }
    }

    private void Redirect(string msg)
    {
        StrFunc.GenErrMsg(msg, string.Format("location.href=\"PublishEdit.aspx?id={0}&pid={1}&topCat={2}&page={3}&pagesize={4}\"", id, pid, topCat, page, pageSize));
    }

    public string getUnit(int articleid, int CtRootId, string CtUnitId)
    {
        if (CtRootId == 4)
        {
            #region //previous version.
            //var result = (from p in _coa_repository.List<CATEGORY>()
            //             where p.DATA_BASE_ID == "DB020" && p.CATEGORY_ID == CtUnitId
            //             select p.CATEGORY_NAME).FirstOrDefault();
            #endregion

            var result = "";// WebCall_GetCategoryDisplayNameById(CtUnitId);
            return (result ?? "");
        }
	    else
        {
            var result = (from p in _mGIPcoanew_repository.List<CuDTGeneric>()
                         join s in _mGIPcoanew_repository.List<CtUnit>() on p.iCTUnit equals s.CtUnitID
                         where p.iCUItem == articleid
                         select s.CtUnitName).FirstOrDefault();

            return (result ?? "");
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
