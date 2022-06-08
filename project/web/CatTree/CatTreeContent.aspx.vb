Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Xml
Partial Class CatTreeContent
    Inherits System.Web.UI.Page

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("KMConnString").ConnectionString
    Dim ConnString2 As String = WebConfigurationManager.ConnectionStrings("LambdaCoaConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim DatabaseId As String = ""
    Dim CategoryId As String = ""
    Dim ReportId As String = ""
    Dim ActorType As String = ""
    Dim ActorId As String = ""
    Dim xmlpath As String = ""
    Protected innerTag As String = ""
    Protected relationsArticle As String = ""
    'Dim Degree As String = ""
    'Dim autoId As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Dim MemberId As String = ""

        DatabaseId = Request("DatabaseId")
        CategoryId = Request("CategoryId")
        ReportId = Request("ReportId")
        ActorType = Request("ActorType")
        If Session("gstyle") Is Nothing Then
            ActorId = ""
        Else
            ActorId = CType(Session("gstyle"), String)
        End If

        If DatabaseId = "" Or CategoryId = "" Or ReportId = "" Then
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        End If
        '---defalut is 消費者---
        If ActorType = "" Then
            ActorType = "002"
        End If
        If ActorType <> "001" And ActorType <> "002" And ActorType <> "003" Then
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        End If

        '---學者---
        If ActorType = "003" Then
            '---檢查---
            If ActorId = "003" Or ActorId = "005" Then
                '---可閱讀---
            Else
                '---無權限---
                '---無權限---
                'alert('您無使用權限')
                Response.Write("<script>window.location.href='/knowledge/';</script>")
                Response.End()
                Exit Sub
            End If
        End If

        '---kpi use---reflash wont add grade---
        If Request.QueryString("kpi") <> "0" Then
            HandleKpiBrowse()
            Dim relink As String = "/CatTree/CatTreeContent.aspx?ReportId=" & ReportId & "&DatabaseId=" & DatabaseId & "&CategoryId=" & CategoryId & "&ActorType=" & ActorType & "&kpi=0"
            Response.Redirect(relink)
            Response.End()
        End If
        '---end of kpi use---

        '加上轉址的處理
        Dim transferUrl As String = "/Category/categorycontent.aspx?CategoryId={0}&ActorType={1}&ReportId={2}"
        If String.IsNullOrEmpty(ReportId) = True Then
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        Else
            '取得對應的CategoryID
            myConnection = New SqlConnection(ConnString2)
            myConnection.Open()
            '查出
            Dim newCategoryID As String = String.Empty

            sqlString = "SELECT DISTINCT NewCategoryId FROM COA_CATEGORY WHERE (DATA_BASE_ID = @dbid) AND (CATEGORY_ID = @catid)"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
            myCommand.Parameters.AddWithValue("@catid", CategoryId)
            myReader = myCommand.ExecuteReader()
            If myReader.FieldCount <> 1 Then
                transferUrl = "/mp.asp?mp=1"
            Else
                If myReader.Read() Then
                    newCategoryID = myReader("NewCategoryId")
                End If
            End If

            If Not myReader.IsClosed Then
                myReader.Close()
            End If

            sqlString = "Select LAMBDA_DocId from COA_REPORT WHERE REPORT_ID = @REPORT_ID "

            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@REPORT_ID", ReportId)
            myReader = myCommand.ExecuteReader()
            If myReader.FieldCount <> 1 Then
                transferUrl = "/mp.asp?mp=1"
            Else
                If myReader.Read() Then
                    If myReader("LAMBDA_DocId") Is DBNull.Value = False AndAlso NOT String.IsNullOrEmpty(myReader("LAMBDA_DocId")) Then
                        transferUrl = String.Format(transferUrl, newCategoryID, ActorType, myReader("LAMBDA_DocId"))
                    Else
                        transferUrl = "/mp.asp?mp=1"
                    End If
                End If
            End If

            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()
            Response.Redirect(transferUrl)
            Response.End()
        End If


        Dim link As String = ""
        link = "CatTreeList.aspx?DatabaseId={dbid}&CategoryId={catid}&ActorType={actortype}"

        If ActorType = "001" Then

            xmlpath = System.Web.Configuration.WebConfigurationManager.AppSettings("CatTreeXmlPath") & "\" & DatabaseId & "\" & ActorType & "\" & DatabaseId & "@" & CategoryId & ".xml"

            NavUrlText.Text = "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", ActorType) & """ title=""生產者知識庫"">生產者知識庫</a>"
            NavTitleText.Text = "生產者知識庫"

        ElseIf ActorType = "002" Then

            xmlpath = System.Web.Configuration.WebConfigurationManager.AppSettings("CatTreeXmlPath") & "\" & DatabaseId & "\" & ActorType & "\" & DatabaseId & "@" & CategoryId & ".xml"

            NavUrlText.Text = "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", ActorType) & """ title=""消費者知識庫"">消費者知識庫</a>"
            NavTitleText.Text = "消費者知識庫"

        ElseIf ActorType = "003" Then

            xmlpath = System.Web.Configuration.WebConfigurationManager.AppSettings("CatTreeXmlPath") & "\" & DatabaseId & "\" & ActorType & "\" & DatabaseId & "@" & CategoryId & ".xml"

            NavUrlText.Text = "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", ActorType) & """ title=""學者知識庫"">學者知識庫</a>"
            NavTitleText.Text = "學者知識庫"

        End If

        '設定左方選單=====================================================================
        If CategoryId = "" Then
            CatTreeViewDS.DataFile = xmlpath
        Else
            Dim sqlstr As String = ""
            sqlstr = sqlstr + " declare @dbId varchar(30)"
            sqlstr = sqlstr + " set @dbId			= 'db020'"
            sqlstr = sqlstr + " "
            sqlstr = sqlstr + " select "
            sqlstr = sqlstr + " 	category_Id, category_name, (select	sub.display_order from category sub where sub.data_base_id = 'db020' and sub.category_id = left(category.category_Id,3) ) as display_order"
            sqlstr = sqlstr + " from category "
            sqlstr = sqlstr + " where "
            sqlstr = sqlstr + " 	data_base_id  = @dbId"
            sqlstr = sqlstr + " and category_id in ('00A' ,'00B' ,'00C' ,'00D' ,'00E' ,'AA0' ,'00G' ,'00H' ,'00I' ,'00J' ,'00K' ,'00L' ,'00M' ,'00N' ,'00Z')"
            sqlstr = sqlstr + " and not category_id like @target_CateId + '%'"
            sqlstr = sqlstr + " union"
            sqlstr = sqlstr + " select"
            sqlstr = sqlstr + " 	category_Id, category_name, (select	sub.display_order from category sub where sub.data_base_id = 'db020' and sub.category_id = left(category.category_Id,3) ) as display_order "
            sqlstr = sqlstr + " from category "
            sqlstr = sqlstr + " where "
            sqlstr = sqlstr + " 	data_base_id  = @dbId"
            sqlstr = sqlstr + " and "

            Select Case Len(CategoryId)
                Case 3, 5
                    sqlstr = sqlstr + " ("
                    sqlstr = sqlstr + " 	category_id  like left(@target_CateId, 3) + REPLICATE ('_', len(@target_CateId) - 3)"
                    sqlstr = sqlstr + " or  category_id  like @target_CateId + '__'"
                    sqlstr = sqlstr + " )"
                Case 7
                    sqlstr = sqlstr + " ("
                    sqlstr = sqlstr + " 	category_id  like left(@target_CateId, 5) + REPLICATE ('_', len(@target_CateId) - 5)"
                    sqlstr = sqlstr + " or  category_id like left(@target_CateId, 3) + '__'"
                    sqlstr = sqlstr + " )"
            End Select
            sqlstr = sqlstr + " order by display_order,  category.category_Id "

            Dim UrlStr As String = "CatTreeList.aspx?DatabaseId=DB020&CategoryId={0}&ActorType={1}"
            Dim doc As XmlDocument = New XmlDocument()
            Dim xmlRoot As XmlNode = doc.CreateNode(XmlNodeType.Element, "node", "")
            Dim xmlAttributeName As XmlAttribute = doc.CreateAttribute("name")
            Dim xmlAttributeUrl As XmlAttribute = doc.CreateAttribute("Url")
            xmlAttributeName.Value = NavTitleText.Text
            xmlAttributeUrl.Value = String.Format(UrlStr, "", ActorType)
            xmlRoot.Attributes.Append(xmlAttributeName)
            xmlRoot.Attributes.Append(xmlAttributeUrl)
            doc.AppendChild(xmlRoot)


            Dim xmlCategoryL1 As XmlNode = doc.CreateNode(XmlNodeType.Element, "node", "")
            Dim xmlCategoryL2 As XmlNode = doc.CreateNode(XmlNodeType.Element, "node", "")
            Dim xmlCategoryL3 As XmlNode = doc.CreateNode(XmlNodeType.Element, "node", "")

            Using myConnection = New SqlConnection(ConnString)
                Using myCommand = New SqlCommand(sqlstr, myConnection)
                    myCommand.Parameters.AddWithValue("@target_CateId", CategoryId)
                    myConnection.Open()
                    myReader = myCommand.ExecuteReader()
                    While myReader.Read()
                        Select Case Len(myReader("category_Id"))
                            Case 3
                                xmlCategoryL1 = doc.CreateNode(XmlNodeType.Element, "node", "")
                                xmlAttributeName = doc.CreateAttribute("name")
                                xmlAttributeUrl = doc.CreateAttribute("Url")
                                xmlAttributeName.Value = myReader("category_Name")
                                xmlAttributeUrl.Value = String.Format(UrlStr, myReader("category_Id"), ActorType)
                                xmlCategoryL1.Attributes.Append(xmlAttributeName)
                                xmlCategoryL1.Attributes.Append(xmlAttributeUrl)
                                xmlRoot.AppendChild(xmlCategoryL1)
                            Case 5
                                xmlCategoryL2 = doc.CreateNode(XmlNodeType.Element, "node", "")
                                xmlAttributeName = doc.CreateAttribute("name")
                                xmlAttributeUrl = doc.CreateAttribute("Url")
                                xmlAttributeName.Value = myReader("category_Name")
                                xmlAttributeUrl.Value = String.Format(UrlStr, myReader("category_Id"), ActorType)
                                xmlCategoryL2.Attributes.Append(xmlAttributeName)
                                xmlCategoryL2.Attributes.Append(xmlAttributeUrl)
                                xmlCategoryL1.AppendChild(xmlCategoryL2)
                            Case 7
                                xmlCategoryL3 = doc.CreateNode(XmlNodeType.Element, "node", "")
                                xmlAttributeName = doc.CreateAttribute("name")
                                xmlAttributeUrl = doc.CreateAttribute("Url")
                                xmlAttributeName.Value = myReader("category_Name")
                                xmlAttributeUrl.Value = String.Format(UrlStr, myReader("category_Id"), ActorType)
                                xmlCategoryL3.Attributes.Append(xmlAttributeName)
                                xmlCategoryL3.Attributes.Append(xmlAttributeUrl)
                                xmlCategoryL2.AppendChild(xmlCategoryL3)
                        End Select
                    End While
                End Using
            End Using

            CatTreeViewDS.Data = doc.InnerXml
        End If

        '設定左方選單end=====================================================================

        '---for navigate use---
        myConnection = New SqlConnection(ConnString)
        Try
            If CategoryId.Length = 3 Then

                sqlString = "SELECT DISTINCT CATEGORY_ID AS CatId1, CATEGORY_NAME AS CatName1 FROM CATEGORY WHERE (CATEGORY.DATA_BASE_ID = @dbid) AND (CATEGORY.CATEGORY_ID = @catid) "
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
                myCommand.Parameters.AddWithValue("@catid", CategoryId)
                myReader = myCommand.ExecuteReader()
                If myReader.Read() Then
                    NavUrlText.Text &= "&gt;" & "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", myReader("CatId1")).Replace("{actortype}", ActorType) & """ title=""" & myReader("CatName1") & """>" & myReader("CatName1") & "</a>"
                End If

            ElseIf CategoryId.Length = 5 Then

                sqlString = "SELECT DISTINCT CATEGORY.CATEGORY_ID AS CatId1, CATEGORY.CATEGORY_NAME AS CatName1, CATEGORY_1.CATEGORY_ID AS CatID2, "
                sqlString &= "CATEGORY_1.CATEGORY_NAME AS CatName2 FROM CATEGORY AS CATEGORY_1 INNER JOIN CATEGORY ON CATEGORY_1.PARENT_CATEGORY_ID = CATEGORY.CATEGORY_ID "
                sqlString &= "WHERE (CATEGORY.DATA_BASE_ID = @dbid) AND (CATEGORY_1.DATA_BASE_ID = @dbid) AND (CATEGORY_1.CATEGORY_ID = @catid) "
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
                myCommand.Parameters.AddWithValue("@catid", CategoryId)
                myReader = myCommand.ExecuteReader()
                If myReader.Read() Then
                    NavUrlText.Text &= "&gt;" & "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", myReader("CatId1")).Replace("{actortype}", ActorType) & """ title=""" & myReader("CatName1") & """>" & myReader("CatName1") & "</a>"
                    NavUrlText.Text &= "&gt;" & "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", myReader("CatId2")).Replace("{actortype}", ActorType) & """ title=""" & myReader("CatName2") & """>" & myReader("CatName2") & "</a>"
                End If

            ElseIf CategoryId.Length = 7 Then

                sqlString = "SELECT DISTINCT CATEGORY.CATEGORY_ID AS CatId1, CATEGORY.CATEGORY_NAME AS CatName1, CATEGORY_1.CATEGORY_ID AS CatID2, "
                sqlString &= "CATEGORY_1.CATEGORY_NAME AS CatName2, CATEGORY_2.CATEGORY_ID AS CatId3, CATEGORY_2.CATEGORY_NAME AS CatName3 "
                sqlString &= "FROM CATEGORY AS CATEGORY_2 INNER JOIN CATEGORY AS CATEGORY_1 ON CATEGORY_2.PARENT_CATEGORY_ID = CATEGORY_1.CATEGORY_ID INNER JOIN "
                sqlString &= "CATEGORY ON CATEGORY_1.PARENT_CATEGORY_ID = CATEGORY.CATEGORY_ID "
                sqlString &= "WHERE (CATEGORY.DATA_BASE_ID = @dbid) AND (CATEGORY_1.DATA_BASE_ID = @dbid) AND (CATEGORY_2.DATA_BASE_ID = @dbid) AND (CATEGORY_2.CATEGORY_ID = @catid) "
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
                myCommand.Parameters.AddWithValue("@catid", CategoryId)
                myReader = myCommand.ExecuteReader()
                If myReader.Read() Then
                    NavUrlText.Text &= "&gt;" & "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", myReader("CatId1")).Replace("{actortype}", ActorType) & """ title=""" & myReader("CatName1") & """>" & myReader("CatName1") & "</a>"
                    NavUrlText.Text &= "&gt;" & "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", myReader("CatId2")).Replace("{actortype}", ActorType) & """ title=""" & myReader("CatName2") & """>" & myReader("CatName2") & "</a>"
                    NavUrlText.Text &= "&gt;" & "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", myReader("CatId3")).Replace("{actortype}", ActorType) & """ title=""" & myReader("CatName3") & """>" & myReader("CatName3") & "</a>"
                End If

            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

        Catch ex As Exception
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

        Dim sb As StringBuilder = New StringBuilder()
        Dim attachCount As Integer = 0
        link = ""
        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT REPORT.SUBJECT, REPORT.CLICK_COUNT, REPORT.AUTHOR, REPORT.DESCRIPTION, REPORT_TYPE1.REPORT_TYPE1_NAME, "
            sqlString &= "REPORT.ONLINE_DATE, REPORT_ATTACH.FILE_NAME, REPORT_ATTACH.FILE_ID, REPORT_ATTACH.FILE_SIZE, REPORT_ATTACH.FILE_DESC "
            sqlString &= "FROM CAT2RPT INNER JOIN REPORT ON CAT2RPT.REPORT_ID = REPORT.REPORT_ID INNER JOIN "
            sqlString &= "CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID INNER JOIN REPORT_TYPE1 ON "
            sqlString &= "REPORT.REPORT_TYPE1_ID = REPORT_TYPE1.REPORT_TYPE1_ID LEFT OUTER JOIN REPORT_ATTACH ON REPORT.REPORT_ID = REPORT_ATTACH.REPORT_ID "
            sqlString &= "INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID "
            sqlString &= "WHERE (CATEGORY.DATA_BASE_ID = @dbid) AND (CATEGORY.CATEGORY_ID = @catid) AND (REPORT.REPORT_ID = @reportid) "
            sqlString &= "AND (RESOURCE_RIGHT.ACTOR_INFO_ID = @actid)"

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
            myCommand.Parameters.AddWithValue("@catid", CategoryId)
            myCommand.Parameters.AddWithValue("@reportid", ReportId)
            If ActorType = "003" Then
                myCommand.Parameters.AddWithValue("@actid", "004")
            Else
                myCommand.Parameters.AddWithValue("@actid", ActorType)
            End If


            myReader = myCommand.ExecuteReader()

            If myReader.Read Then


                If Session("BrowseTitleNew") <> myReader("SUBJECT") Then
                    Session("BrowseTitleNew") = myReader("SUBJECT")
                    Session("BrowseLinkNew") = "/CatTree/CatTreeContent.aspx?ReportId=" & ReportId & "&DatabaseId=" & DatabaseId & "&CategoryId=" & CategoryId & "&ActorType=" & ActorType
                End If

                sb.AppendLine("<ul class=""Function2"" xmlns:hyweb=""urn:gip-hyweb-com"" xmlns="""">")
                sb.AppendLine(GetCommendWordBtn())
                'added by Joey, 新增問題反應
                'http://gssjira.gss.com.tw/browse/COAKM-30，Domain 不同造成showModalDialog失效
                'Response.Write(me.request.url.tostring)
                sb.AppendLine("<li><a href=""#"" class=""Rword"" title=""系統問題"" onclick=""window.showModalDialog('http://" & Request.Url.Authority & "/mailbox.asp?a=a9191',self);return false;"">")
                sb.AppendLine("系統問題</a></li>")
                sb.Append("<input type=""hidden""  name=""type"" value=""3"" />")
                sb.Append("<input type=""hidden""  name=""ARTICLE_ID"" value=" & ReportId & " />")

                sb.AppendLine("<li><a class=""Print"" onclick=""window.open('CatTreePrintContent.aspx" + Request.Url.Query + "')"" href='#' title=""友善列印"" >")
                sb.AppendLine("友善列印</a></li>")
                sb.AppendLine("<li><a href=""javascript:history.go(-1);"" class=""Back"" title=""回上一頁"">回上一頁</a><noscript>	本網頁使用SCRIPT編碼方式執行回上一頁的動作，如果您的瀏覽器不支援SCRIPT，請直接使用「Alt+方向鍵向左按鍵」來返回上一頁</noscript></li></ul>")
                sb.AppendLine("<div id=""cp"" xmlns="""">")
                sb.AppendLine("<table summary=""排版表格"" class=""cptable"">")
                sb.AppendLine("<col class=""title"">")
                sb.AppendLine("<col class=""cptablecol2"">")
                sb.AppendLine("<tr><th scope=""row"">標題</th><td>" & myReader("SUBJECT") & "</td></tr>")
                sb.AppendLine("<tr><th scope=""row"">內文</th><td>" & ReplaceAndFindKeyword(myReader("DESCRIPTION").ToString().Replace(vbCrLf, "<br/>")) & "</td></tr>")
                sb.AppendLine("<tr><th scope=""row"">文件屬性</th><td>" & myReader("REPORT_TYPE1_NAME") & "</td></tr>")
                sb.AppendLine("<tr><th scope=""row"">點閱次數</th><td>" & myReader("CLICK_COUNT") & "</td></tr>")
                sb.AppendLine("<tr><th scope=""row"">作者</th><td>" & myReader("AUTHOR") & "</td></tr>")
                sb.AppendLine("<tr><th scope=""row"">知識樹分類</th><td>" & ListCategory() & "</td></tr>")
                sb.AppendLine("<tr><th scope=""row"">張貼日期</th><td>" & myReader("ONLINE_DATE") & "</td></tr>")
                sb.AppendLine("</table>")
                If Not myReader("FILE_ID").ToString = "" Then
                    If ActorType = "003" Then
                        link = "http://kmintra.coa.gov.tw/coa/InternetDownloadFileServlet?file_id=" & myReader("FILE_ID") & "&actor_info_id=004"
                    Else
                        link = "http://kmintra.coa.gov.tw/coa/InternetDownloadFileServlet?file_id=" & myReader("FILE_ID") & "&actor_info_id=" & ActorType
                    End If

                    sb.AppendLine("<div class=""download""><h5>相關檔案下載</h5>")
                    sb.AppendLine("<ul><li><a target=""_blank"" href=""" & link & """ title=""" & myReader("FILE_DESC") & """>" & myReader("FILE_NAME") & "</a> (" & myReader("FILE_SIZE") & " KB) </li></ul>")
                    attachCount += 1
                    While myReader.Read
                        If ActorType = "003" Then
                            link = "http://kmintra.coa.gov.tw/coa/InternetDownloadFileServlet?file_id=" & myReader("FILE_ID") & "&actor_info_id=004"
                        Else
                            link = "http://kmintra.coa.gov.tw/coa/InternetDownloadFileServlet?file_id=" & myReader("FILE_ID") & "&actor_info_id=" & ActorType
                        End If
                        sb.AppendLine("<ul><li><a target=""_blank"" href=""" & link & """ title=""" & myReader("FILE_DESC") & """>" & myReader("FILE_NAME") & "</a>(" & myReader("FILE_SIZE") & " KB)</li>  </ul>")
                        attachCount += 1
                    End While
                    sb.AppendLine("</div>")

                End If
                sb.AppendLine("</div>")
            Else
                Response.Write("<script>alert('您無權限閱讀此文章');history.back();</script>")
                Exit Sub
            End If

            TableText.Text = sb.ToString()

            '-延伸閱讀開始-'
            Dim reloationsIndex As Integer = 1
            Dim reloationStr As String = ""
            myConnection = New SqlConnection(ConnString)
            Try
                Dim sqlStr As StringBuilder = New StringBuilder()
                sqlStr.Append(" select top(5) category_id,coa.dbo.report.report_id ,coa.dbo.report.Subject,coa.dbo.report.create_date, " & vbCrLf)
                sqlStr.Append(" subTable.CNT  " & vbCrLf)
                sqlStr.Append(" from coa.dbo.report  " & vbCrLf)
                sqlStr.Append(" inner join  " & vbCrLf)
                sqlStr.Append(" (  " & vbCrLf)
                sqlStr.Append(" 	SELECT RK.REPORT_ID,CNT,CAT.CATEGORY_ID FROM ( " & vbCrLf)
                sqlStr.Append(" 		select REPORT_ID,count(*) as CNT " & vbCrLf)
                sqlStr.Append(" 			FROM coa.dbo.REPORT_KEYWORD_FREQUENCY  " & vbCrLf)
                sqlStr.Append(" 			where REPORT_ID != @REPORT_ID and Keyword in  " & vbCrLf)
                sqlStr.Append(" 			( " & vbCrLf)
                sqlStr.Append(" 				SELECT Keyword  " & vbCrLf)
                sqlStr.Append(" 				FROM coa.dbo.REPORT_KEYWORD_FREQUENCY  " & vbCrLf)
                sqlStr.Append(" 				where REPORT_ID = @REPORT_ID " & vbCrLf)
                sqlStr.Append(" 				AND NOT KEYWORD IN ( " & vbCrLf)
                sqlStr.Append(" 					select TOP 20 KEYWORD " & vbCrLf)
                sqlStr.Append(" 					from coa.dbo.REPORT_KEYWORD_FREQUENCY " & vbCrLf)
                sqlStr.Append(" 					GROUP BY KEYWORD " & vbCrLf)
                sqlStr.Append(" 					ORDER BY COUNT(*) DESC	 " & vbCrLf)
                sqlStr.Append(" 				)  " & vbCrLf)
                sqlStr.Append(" 			) " & vbCrLf)
                sqlStr.Append(" 		group by report_Id " & vbCrLf)
                sqlStr.Append(" 	) AS RK " & vbCrLf)
                sqlStr.Append(" 	LEFT JOIN ( " & vbCrLf)
                sqlStr.Append(" 		SELECT * FROM COA..CAT2RPT A " & vbCrLf)
                sqlStr.Append(" 		WHERE DATA_BASE_ID= @DatabaseId " & vbCrLf)
                sqlStr.Append(" 		AND CATEGORY_ID = ( " & vbCrLf)
                sqlStr.Append(" 		SELECT MAX(CATEGORY_ID)  " & vbCrLf)
                sqlStr.Append(" 		FROM COA..CAT2RPT B " & vbCrLf)
                sqlStr.Append(" 		WHERE B.REPORT_ID=A.REPORT_ID AND DATA_BASE_ID='DB020' " & vbCrLf)
                sqlStr.Append(" 		)	 " & vbCrLf)
                sqlStr.Append(" 	) CAT  " & vbCrLf)
                sqlStr.Append(" 	on CAT.report_id = RK.report_id " & vbCrLf)
                sqlStr.Append(" 	LEFT JOIN coa.dbo.RESOURCE_RIGHT  " & vbCrLf)
                sqlStr.Append(" 	on actor_info_id = '002' and resource_id = 'report@' + RK.report_id  " & vbCrLf)
                sqlStr.Append(" 	WHERE NOT resource_id IS NULL AND NOT CAT.data_base_id IS NULL " & vbCrLf)
                sqlStr.Append(" )  " & vbCrLf)
                sqlStr.Append(" subTable  " & vbCrLf)
                sqlStr.Append(" on subTable.report_id = coa.dbo.report.report_id  " & vbCrLf)
                sqlStr.Append(" order by subTable.CNT desc,coa.dbo.report.create_date desc " & vbCrLf)
                myConnection.Open()
                myCommand = New SqlCommand(sqlStr.ToString(), myConnection)
                myCommand.Parameters.Add(New SqlParameter("@REPORT_ID", ReportId))
                myCommand.Parameters.Add(New SqlParameter("@DatabaseId", DatabaseId))
                myReader = myCommand.ExecuteReader()
                While myReader.Read
                    If reloationsIndex = 1 Then
                        reloationStr &= "<table class=""similar"">"
                        reloationStr &= "<tr><th colSpan=""3"">延伸閱讀</th></tr>"
                    End If
                    reloationStr &= "<tr><td width=""6%"">" & reloationsIndex & "</td>"
                    reloationStr &= "<td width=""74%""><a href=""/CatTree/CatTreeContent.aspx?ReportId=" & myReader("report_id") & "&DatabaseId=" & DatabaseId & "&CategoryId=" & myReader("category_id") & """>" & myReader("Subject") & "</a></td>"
                    reloationStr &= "<td width=""20%"">[" & Date.Parse(myReader("create_date")).ToShortDateString & "]</td></tr>"
                    reloationsIndex = reloationsIndex + 1
                End While
                If reloationStr <> "" Then
                    reloationStr &= "</table><BR/>"
                End If
                myCommand.Dispose()
            Catch ex As Exception
                Response.Write(ex.ToString())
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try
            relationsArticle = reloationStr
            '-延伸閱讀結束-'

            '---for report use---
            myConnection = New SqlConnection(ConnString)
            Try
                sqlString = "UPDATE REPORT SET CLICK_COUNT = CLICK_COUNT + 1 WHERE REPORT_ID = @reportid"
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.Add(New SqlParameter("@reportid", ReportId))
                myCommand.ExecuteNonQuery()
                myCommand.Dispose()
            Catch ex As Exception
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try

        Catch ex As Exception
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Sub

    Public Function ListCategory()

        Dim sb As New StringBuilder
        Dim myReader1 As SqlDataReader
        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT DISTINCT CATEGORY.CATEGORY_ID AS CatId1, CATEGORY.CATEGORY_NAME AS CatName1, "
            sqlString &= "CATEGORY_1.CATEGORY_ID as CatId2, CATEGORY_1.CATEGORY_NAME AS CatName2, "
            sqlString &= "CATEGORY_2.CATEGORY_ID as CatId3, CATEGORY_2.CATEGORY_NAME AS CatName3 "
            sqlString &= "FROM CATEGORY AS CATEGORY_2"
            sqlString &= " RIGHT OUTER JOIN CATEGORY AS CATEGORY_1 ON CATEGORY_2.CATEGORY_ID = CATEGORY_1.PARENT_CATEGORY_ID "
            sqlString &= " RIGHT OUTER JOIN REPORT INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID "
            sqlString &= " INNER JOIN  CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID ON CATEGORY_1.CATEGORY_ID = CATEGORY.PARENT_CATEGORY_ID "
            sqlString &= " WHERE "
            sqlString &= "     (CATEGORY.DATA_BASE_ID = @dbid) "
            sqlString &= " and (CATEGORY_1.DATA_BASE_ID = @dbid) "
            sqlString &= " and (CAT2RPT.DATA_BASE_ID = @dbid) "
            sqlString &= " AND (REPORT.REPORT_ID = @reportid) "
            sqlString &= " ORDER BY CatId1"

            sqlString = " Select DISTINCT"
            sqlString &= " CATEGORY.CATEGORY_ID AS CatId1, CATEGORY.CATEGORY_NAME AS CatName1"
            sqlString &= " , CATEGORY_1.CATEGORY_ID AS CatId2"
            sqlString &= " , CATEGORY_1.CATEGORY_NAME AS CatName2"
            sqlString &= " , CATEGORY_2.CATEGORY_ID AS CatId3"
            sqlString &= " , CATEGORY_2.CATEGORY_NAME AS CatName3"
            sqlString &= " FROM REPORT "
            sqlString &= "      INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID "
            sqlString &= "      INNER JOIN CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID AND CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID "
            sqlString &= "      left  JOIN CATEGORY AS CATEGORY_1 ON CATEGORY.PARENT_CATEGORY_ID = CATEGORY_1.CATEGORY_ID AND CATEGORY.DATA_BASE_ID = CATEGORY_1.DATA_BASE_ID "
            sqlString &= "      left  JOIN CATEGORY AS CATEGORY_2 ON CATEGORY_1.PARENT_CATEGORY_ID = CATEGORY_2.CATEGORY_ID AND CATEGORY_1.DATA_BASE_ID = CATEGORY_2.DATA_BASE_ID"
            sqlString &= "  WHERE"
            sqlString &= " 	    (CAT2RPT.DATA_BASE_ID = @dbid) "
            sqlString &= "  AND (REPORT.REPORT_ID = @reportid)"
            sqlString &= " ORDER BY CatId1"


            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
            myCommand.Parameters.AddWithValue("@reportid", ReportId)
            myReader1 = myCommand.ExecuteReader()

            sb.AppendLine("<ul>")
            While myReader1.Read()

                sb.AppendLine("<li>" & NavTitleText.Text)

                If myReader1.Item("CatId3") IsNot DBNull.Value Then
                    sb.AppendLine("&gt; " & myReader1("CatName3"))
                End If
                If myReader1.Item("CatId2") IsNot DBNull.Value Then
                    sb.AppendLine("&gt; " & myReader1("CatName2"))
                End If
                If myReader1.Item("CatId1") IsNot DBNull.Value Then
                    sb.AppendLine("&gt; " & myReader1("CatName1"))
                End If
                sb.AppendLine("</li>")

            End While
            sb.AppendLine("</ul>")

            If Not myReader1.IsClosed Then
                myReader1.Close()
            End If
            myCommand.Dispose()

        Catch ex As Exception
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

        Return sb.ToString

    End Function

    Function GetCommendWordBtn() As String

        Dim str As String = ""

        Dim sb As New StringBuilder
        Dim ConnStr As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim myReader1 As SqlDataReader
        myConnection = New SqlConnection(ConnStr)
        Try
            sqlString = "SELECT codeMetaID, mCode, mValue, mSortValue FROM CodeMain WHERE (codeMetaID = 'commendWord') AND (mCode = 'tank')"

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myReader1 = myCommand.ExecuteReader()

            While myReader1.Read()

                If myReader1("mValue") = "1" Then
                    'ReportId=935977&DatabaseId=DB020&CategoryId=00L0401&ActorType=002
                    sb.AppendLine("<li><a href=""javascript:getSelectedText('ReportId=" & Request.QueryString("ReportId") & "&DatabaseId=" & _
                                  Request.QueryString("DatabaseId") & "&CategoryId=" & Request.QueryString("CategoryId") & _
                                  "&ActorType=" & Request.QueryString("ActorType") & "')"" class=""Rword"">推薦詞彙</a></li>")
                End If

            End While

            If Not myReader1.IsClosed Then
                myReader1.Close()
            End If
            myCommand.Dispose()

        Catch ex As Exception
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

        Return sb.ToString

    End Function

    Sub HandleKpiBrowse()

        '---start of kpi user---20080911---vincent---

        Dim sb As StringBuilder = New StringBuilder()
        Dim attachCount As Integer = 0
        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT REPORT_ATTACH.FILE_NAME FROM CAT2RPT INNER JOIN REPORT ON CAT2RPT.REPORT_ID = REPORT.REPORT_ID INNER JOIN "
            sqlString &= "CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID INNER JOIN REPORT_TYPE1 ON "
            sqlString &= "REPORT.REPORT_TYPE1_ID = REPORT_TYPE1.REPORT_TYPE1_ID LEFT OUTER JOIN REPORT_ATTACH ON REPORT.REPORT_ID = REPORT_ATTACH.REPORT_ID "
            sqlString &= "INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID "
            sqlString &= "WHERE (CATEGORY.DATA_BASE_ID = @dbid) AND (CATEGORY.CATEGORY_ID = @catid) AND (REPORT.REPORT_ID = @reportid) "
            sqlString &= "AND (RESOURCE_RIGHT.ACTOR_INFO_ID = @actid)"

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
            myCommand.Parameters.AddWithValue("@catid", CategoryId)
            myCommand.Parameters.AddWithValue("@reportid", ReportId)
            If ActorType = "003" Then
                myCommand.Parameters.AddWithValue("@actid", "004")
            Else
                myCommand.Parameters.AddWithValue("@actid", ActorType)
            End If
            myReader = myCommand.ExecuteReader()
            While myReader.Read
                attachCount += 1
            End While
            If Not myReader.IsClosed Then myReader.Close()
            myCommand.Dispose()
        Catch
        Finally
            If myConnection.State = ConnectionState.Open Then myConnection.Close()
        End Try

        If Session("memID") IsNot Nothing Or Session("memID") <> "" Then
            Dim browse As New KPIBrowse(Session("memID"), "browseCatTreeCP", attachCount.ToString)
            browse.HandleBrowse()
        End If

    End Sub

    Function ReplaceAndFindKeyword(ByVal strTemp As String) As String
        strTemp = strTemp.Replace(Chr(13), "<br />")
        Dim string1
        string1 = PediaUtility.ReplacePedia(strTemp)
        strTemp = string1(0)
        innerTag &= string1(1)
        Return strTemp
    End Function

End Class
