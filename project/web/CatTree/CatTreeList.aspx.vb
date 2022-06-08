Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Xml
Partial Class CatTreeList
    Inherits System.Web.UI.Page

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("KMConnString").ConnectionString
    Dim ConnString2 As String = WebConfigurationManager.ConnectionStrings("LambdaCoaConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim myTable As DataTable
    Dim myDataRow As DataRow
    Dim DatabaseId As String = ""
    Dim CategoryId As String = ""
    Dim ActorType As String = ""
    Dim ActorId As String = ""
    Dim xmlpath As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        DatabaseId = Request("DatabaseId")
        CategoryId = Request("CategoryId")
        ActorType = Request("ActorType")
        CatTreeViewDS.Data = ""
        If Request.RawUrl.Contains("%3c") Or Request.RawUrl.Contains("%3d") Or Request.RawUrl.Contains("%3e") Or _
           Request.RawUrl.Contains("%22") Or Request.RawUrl.Contains("%") Then
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        End If
        If DatabaseId = "" Then
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        Else
            If DatabaseId <> "DB020" Then
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            End If
        End If
        If CategoryId <> "" Then
            If Not CategoryId.StartsWith("00") And Not CategoryId.StartsWith("AA0") Then
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            End If
        End If
        If Request.QueryString("Browse") <> "" And Request.QueryString("Browse") <> "clear" Then
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        End If
        If Request.QueryString("PageSize") <> "" Then
            Try
                Dim int As Integer = Integer.Parse(Request.QueryString("PageSize"))
            Catch ex As Exception
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            End Try
        End If
        If Request.QueryString("PageNumber") <> "" Then
            Try
                Dim int As Integer = Integer.Parse(Request.QueryString("PageNumber"))
            Catch ex As Exception
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            End Try
        End If
        If Request.QueryString("t") <> "" And Request.QueryString("t") <> "a" And Request.QueryString("t") <> "b" Then
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        End If

        If Session("gstyle") Is Nothing Then
            ActorId = ""
        Else
            ActorId = CType(Session("gstyle"), String)
        End If

        '---defalut is 知識庫首頁---
        If ActorType = "" Then
            'ActorType = "002"
            ActorType = "000"
        End If

        If ActorType <> "000" And ActorType <> "001" And ActorType <> "002" And ActorType <> "003" Then
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
                'alert('您無使用權限')
                Response.Write("<script>alert('限學者會員方可進入!');history.go(-1);;</script>")
                Response.End()
            End If
        End If

        '加上轉址的處理
        Dim transferUrl As String = "/Category/categorylist.aspx?CategoryId={0}&ActorType={1}"
        If String.IsNullOrEmpty(CategoryId) = True Then
            If (Request.QueryString("t") <> "") Then
                transferUrl += "&t=" & Request.QueryString("t")
            End If
            Response.Redirect(String.Format(transferUrl, String.Empty, ActorType))
            Response.End()
        Else
            '取得對應的CategoryID
            myConnection = New SqlConnection(ConnString2)

            sqlString = "SELECT DISTINCT NewCategoryId FROM COA_CATEGORY WHERE (DATA_BASE_ID = @dbid) AND (CATEGORY_ID = @catid) "
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
            myCommand.Parameters.AddWithValue("@catid", CategoryId)
            myReader = myCommand.ExecuteReader()
            If myReader.FieldCount <> 1 Then
                transferUrl = String.Format(transferUrl, String.Empty, ActorType)
            Else
                If myReader.Read() Then
                    transferUrl = String.Format(transferUrl, myReader("NewCategoryId"), ActorType)
                End If
            End If

            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()
            Response.Redirect(transferUrl)
            Response.End()
        End If

        '---tab---
        Dim link As String = ""
        link = "CatTreeList.aspx?DatabaseId={dbid}&CategoryId={catid}&ActorType={actortype}&t=a"

        TabText.Text = "<ul class=""group"">"

        If ActorType = "000" Then

            xmlpath = System.Web.Configuration.WebConfigurationManager.AppSettings("CatTreeXmlPath") & "\" & DatabaseId & "\" & ActorType & "\" & DatabaseId & "@" & CategoryId & ".xml"

            NavUrlText.Text = "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", ActorType) & """ title=""知識庫首頁"">知識庫首頁</a>"
            NavTitleText.Text = "知識庫首頁"

            TabText.Text &= "<li class=""active""><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "000") & """>知識庫首頁</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "002") & """>消費者</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "001") & """>生產者</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "003") & """>學者</a></li>"

        ElseIf ActorType = "001" Then

            xmlpath = System.Web.Configuration.WebConfigurationManager.AppSettings("CatTreeXmlPath") & "\" & DatabaseId & "\" & ActorType & "\" & DatabaseId & "@" & CategoryId & ".xml"

            NavUrlText.Text = "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", ActorType) & """ title=""生產者知識庫"">生產者知識庫</a>"
            NavTitleText.Text = "生產者知識庫"

            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "000") & """>知識庫首頁</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "002") & """>消費者</a></li>"
            TabText.Text &= "<li class=""active""><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "001") & """>生產者</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "003") & """>學者</a></li>"

        ElseIf ActorType = "002" Then

            xmlpath = System.Web.Configuration.WebConfigurationManager.AppSettings("CatTreeXmlPath") & "\" & DatabaseId & "\" & ActorType & "\" & DatabaseId & "@" & CategoryId & ".xml"

            NavUrlText.Text = "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", ActorType) & """ title=""消費者知識庫"">消費者知識庫</a>"
            NavTitleText.Text = "消費者知識庫"

            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "000") & """>知識庫首頁</a></li>"
            TabText.Text &= "<li class=""active""><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "002") & """>消費者</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "001") & """>生產者</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "003") & """>學者</a></li>"

        ElseIf ActorType = "003" Then

            xmlpath = System.Web.Configuration.WebConfigurationManager.AppSettings("CatTreeXmlPath") & "\" & DatabaseId & "\" & ActorType & "\" & DatabaseId & "@" & CategoryId & ".xml"

            NavUrlText.Text = "<a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", ActorType) & """ title=""學者知識庫"">學者知識庫</a>"
            NavTitleText.Text = "學者知識庫"

            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "000") & """>知識庫首頁</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "002") & """>消費者</a></li>"
            TabText.Text &= "<li><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "001") & """>生產者</a></li>"
            TabText.Text &= "<li class=""active""><a href=""" & link.Replace("{dbid}", DatabaseId).Replace("{catid}", "").Replace("{actortype}", "003") & """>學者</a></li>"

        End If

        TabText.Text &= "</ul>"

        '設定左方選單=====================================================================
        If CategoryId = "" Or (Len(CategoryId) <> 3 And Len(CategoryId) <> 5 And Len(CategoryId) <> 7) Then
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
                    myConnection.Open()
                    myCommand.Parameters.AddWithValue("@target_CateId", CategoryId)
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

        If ActorType <> "000" Then
            '---navigate---
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

                    If Not myReader.IsClosed Then
                        myReader.Close()
                    End If
                    myCommand.Dispose()

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
                    If Not myReader.IsClosed Then
                        myReader.Close()
                    End If
                    myCommand.Dispose()

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
                    If Not myReader.IsClosed Then
                        myReader.Close()
                    End If
                    myCommand.Dispose()
                End If

            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try
        End If

        Dim sbb As New StringBuilder
        Dim tempurl As String = "CatTreeList.aspx?DatabaseId={2}&CategoryId={0}&ActorType={1}&t=a".Replace("{2}", DatabaseId).Replace("{1}", ActorType)
        Dim tempnode As String = "<a href=""{0}"">{1}({2})</a>｜"

        If ActorType <> "000" Then

            sbb.Append("<div id=""dividetype"">子目錄資料數：")
            Try
                sqlString = "SELECT * FROM CAT2RPTCOUNT WHERE DatabaseId = @databaseid AND CategoryId LIKE '" & CategoryId & "%' AND Degree = @degree ORDER BY DisplayOrder"

                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@databaseid", DatabaseId)
                If CategoryId.Length = 0 Then
                    myCommand.Parameters.AddWithValue("@degree", 1)
                ElseIf CategoryId.Length = 3 Then
                    myCommand.Parameters.AddWithValue("@degree", 2)
                ElseIf CategoryId.Length = 5 Then
                    myCommand.Parameters.AddWithValue("@degree", 3)
                ElseIf CategoryId.Length = 7 Then
                    myCommand.Parameters.AddWithValue("@degree", 3)
                End If
                myReader = myCommand.ExecuteReader()

                While myReader.Read()
                    If ActorType = "001" Then
                        sbb.Append(tempnode.Replace("{0}", tempurl.Replace("{0}", myReader("CategoryId"))).Replace("{1}", myReader("CategoryName")).Replace("{2}", myReader("ProducerCount")))
                    ElseIf ActorType = "002" Then
                        sbb.Append(tempnode.Replace("{0}", tempurl.Replace("{0}", myReader("CategoryId"))).Replace("{1}", myReader("CategoryName")).Replace("{2}", myReader("ConsumerCount")))
                    ElseIf ActorType = "003" Then
                        sbb.Append(tempnode.Replace("{0}", tempurl.Replace("{0}", myReader("CategoryId"))).Replace("{1}", myReader("CategoryName")).Replace("{2}", myReader("ScholarCount")))
                    End If
                End While

                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try

            sbb.Append("</div>")
        Else
            sbb.Append("<div class=""nodata"">請選擇上方「消費者」、「生產者」、「學者」分眾知識樹開始瀏覽，其中學者知識樹需具有學者會員資格。選擇分眾知識樹後，您可直接點選左欄欲瀏覽之知識庫節點。</div>")
            sbb.Append("<div class=""hotessay""><h3>熱門知識庫文章</h3><div id=""MagTabs""><ul>")

            If Request("t") = "a" Then
                sbb.Append("<li class=""current""><a href=""" & tempurl.Replace("{0}", "") & """><span>最新文章</span></a></li>")
                sbb.Append("<li><a href=""" & tempurl.Replace("&t=a", "&t=b").Replace("{0}", "") & """><span>最多瀏覽</span></a></li>")
            Else
                sbb.Append("<li><a href=""" & tempurl.Replace("{0}", "") & """><span>最新文章</span></a></li>")
                sbb.Append("<li class=""current""><a href=""" & tempurl.Replace("&t=a", "&t=b").Replace("{0}", "") & """><span>最多瀏覽</span></a></li>")
            End If

            sbb.Append("</ul></div></div>")
        End If

        NodeText.Text = sbb.ToString

        Dim PageSize As Integer = 10
        Dim PageNumber As Integer = 0
        If Not Page.IsPostBack Then
            Try
                If Request("PageSize") = "" Then
                    PageSize = 10
                Else
                    PageSize = CInt(Request("PageSize"))
                End If
                If Request("PageNumber") = "" Then
                    PageNumber = 1
                Else
                    PageNumber = CInt(Request("PageNumber"))
                End If
            Catch ex As Exception
                Response.Redirect("/mp.asp?mp=1")
                Exit Sub
            End Try
        Else
            PageSize = PageSizeDDL.SelectedValue
            PageNumber = PageNumberDDL.SelectedValue
        End If

        Dim total As Integer = 0
        Dim pageCount As Integer = 0
        Dim position As Integer = 1

        Dim sb As StringBuilder = New StringBuilder()
        myConnection = New SqlConnection(ConnString)

        If CategoryId <> "" Then

            Try
                sqlString = " SELECT CAT2RPT.DATA_BASE_ID, CAT2RPT.CATEGORY_ID, CAT2RPT.REPORT_ID, REPORT.SUBJECT, REPORT.CLICK_COUNT, "
                sqlString &= " REPORT_TYPE1.REPORT_TYPE1_NAME, REPORT.ONLINE_DATE "
                sqlString &= " FROM CAT2RPT "
                sqlString &= "  INNER JOIN REPORT ON CAT2RPT.REPORT_ID = REPORT.REPORT_ID "
                sqlString &= "  INNER JOIN CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID  and CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID "
                sqlString &= "  INNER JOIN REPORT_TYPE1 ON REPORT.REPORT_TYPE1_ID = REPORT_TYPE1.REPORT_TYPE1_ID "
                sqlString &= "  INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID "
                sqlString &= " WHERE (CATEGORY.DATA_BASE_ID = @dbid) "
                sqlString &= " AND (CATEGORY.CATEGORY_ID = @catid) "
                sqlString &= " AND (RESOURCE_RIGHT.ACTOR_INFO_ID = @actid)"
                sqlString &= " ORDER BY "
                sqlString &= " REPORT.ONLINE_DATE DESC "

                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@dbid", DatabaseId)
                myCommand.Parameters.AddWithValue("@catid", CategoryId)
                If ActorType = "003" Then
                    myCommand.Parameters.AddWithValue("@actid", "004")
                Else
                    myCommand.Parameters.AddWithValue("@actid", ActorType)
                End If

                myReader = myCommand.ExecuteReader()

                myTable = New DataTable()
                myTable.Load(myReader)

                total = myTable.Rows().Count()
                If total <> 0 Then
                    pageCount = Int(total / PageSize + 0.999)

                    If pageCount < PageNumber Then
                        PageNumber = pageCount
                    End If

                    PageNumberText.Text = PageNumber.ToString()
                    TotalPageText.Text = pageCount.ToString()
                    TotalRecordText.Text = total.ToString()
                    If PageSize = 10 Then
                        PageSizeDDL.SelectedIndex = 0
                    ElseIf PageSize = 20 Then
                        PageSizeDDL.SelectedIndex = 1
                    ElseIf PageSize = 30 Then
                        PageSizeDDL.SelectedIndex = 2
                    ElseIf PageSize = 50 Then
                        PageSizeDDL.SelectedIndex = 3
                    End If
                    Dim item As ListItem
                    Dim j As Integer = 0
                    PageNumberDDL.Items.Clear()
                    For j = 0 To pageCount - 1
                        item = New ListItem
                        item.Value = j + 1
                        item.Text = j + 1
                        If PageNumber = (j + 1) Then
                            item.Selected = True
                        End If
                        PageNumberDDL.Items.Insert(j, item)
                        item = Nothing
                    Next

                    position = PageSize * (PageNumber - 1)

                    myDataRow = myTable.Rows.Item(position)

                    If PageNumber > 1 Then
                        PreviousText.Visible = True
                        PreviousImg.Visible = True
                        PreviousLink.NavigateUrl = "CatTreeList.aspx?DatabaseId=" & DatabaseId & "&CategoryId=" & CategoryId & "&ActorType=" & ActorType & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                    Else
                        PreviousText.Visible = False
                        PreviousImg.Visible = False
                    End If
                    If PageNumberDDL.SelectedValue < pageCount Then
                        NextText.Visible = True
                        NextImg.Visible = True
                        NextLink.NavigateUrl = "CatTreeList.aspx?DatabaseId=" & DatabaseId & "&CategoryId=" & CategoryId & "&ActorType=" & ActorType & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                    Else
                        NextText.Visible = False
                        NextImg.Visible = False
                    End If

                    Dim i As Integer = 0
                    link = ""
                    sb.Append("<div class=""list"">")
                    sb.Append("<ul>")
                    For i = 0 To PageSize - 1

                        link = ""
                        sb.Append("<li>")

                        link = "<a href=""CatTreeContent.aspx?ReportId=" & myDataRow("REPORT_ID") & "&DatabaseId=" & DatabaseId
                        link &= "&CategoryId=" & CategoryId & "&ActorType=" & ActorType & """>" & myDataRow("SUBJECT") & "</a>"

                        sb.Append("<span>" & link & "</span>")
                        sb.Append("<span>" & myDataRow("CLICK_COUNT") & "</span>")
                        sb.Append("<span>" & myDataRow("REPORT_TYPE1_NAME") & "</span>")
                        sb.Append("<span>" & myDataRow("ONLINE_DATE") & "</span>")
                        sb.Append("</li>")
                        position += 1
                        If myTable.Rows.Count <= position Then
                            Exit For
                        End If
                        myDataRow = myTable.Rows.Item(position)
                    Next
                    sb.Append("</ul>")
                    sb.Append("</div>")

                    TableText.Text = sb.ToString()
                Else
                    TableText.Text = "<div class=""nodata"">本層節點目前無資料，您可點選上方子目錄資料項目、或點選左欄其他知識庫節點。</div>"
                End If

            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try

        ElseIf ActorType = "000" Then

            Dim InSql As String = ""

            Try
                sqlString = "SELECT DISTINCT REPORT.REPORT_ID, REPORT.SUBJECT, REPORT.ONLINE_DATE, REPORT.CLICK_COUNT FROM REPORT INNER JOIN "
                sqlString &= "CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID "
                sqlString &= "AND CAT2RPT.DATA_BASE_ID = CATEGORY.DATA_BASE_ID INNER JOIN RESOURCE_RIGHT ON 'report@' + REPORT.REPORT_ID = RESOURCE_RIGHT.RESOURCE_ID "
                sqlString &= "WHERE (CAT2RPT.DATA_BASE_ID = 'DB020') AND (RESOURCE_RIGHT.ACTOR_INFO_ID = '002') "

                If Request("t") = "a" Then
                    sqlString &= "ORDER BY REPORT.ONLINE_DATE DESC"
                ElseIf Request("t") = "b" Then
                    sqlString &= "ORDER BY REPORT.CLICK_COUNT DESC"
                End If
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myReader = myCommand.ExecuteReader()

                myTable = New DataTable()
                myTable.Load(myReader)

                If Not myReader.IsClosed Then
                    myReader.Close()
                End If

                total = myTable.Rows().Count()
                If total <> 0 Then
                    pageCount = Int(total / PageSize + 0.999)

                    If pageCount < PageNumber Then
                        PageNumber = pageCount
                    End If

                    PageNumberText.Text = PageNumber.ToString()
                    TotalPageText.Text = pageCount.ToString()
                    TotalRecordText.Text = total.ToString()
                    If PageSize = 10 Then
                        PageSizeDDL.SelectedIndex = 0
                    ElseIf PageSize = 20 Then
                        PageSizeDDL.SelectedIndex = 1
                    ElseIf PageSize = 30 Then
                        PageSizeDDL.SelectedIndex = 2
                    ElseIf PageSize = 50 Then
                        PageSizeDDL.SelectedIndex = 3
                    End If
                    Dim item As ListItem
                    Dim j As Integer = 0
                    PageNumberDDL.Items.Clear()
                    For j = 0 To pageCount - 1
                        item = New ListItem
                        item.Value = j + 1
                        item.Text = j + 1
                        If PageNumber = (j + 1) Then
                            item.Selected = True
                        End If
                        PageNumberDDL.Items.Insert(j, item)
                        item = Nothing
                    Next

                    position = PageSize * (PageNumber - 1)

                    myDataRow = myTable.Rows.Item(position)

                    If PageNumber > 1 Then
                        PreviousText.Visible = True
                        PreviousImg.Visible = True
                        PreviousLink.NavigateUrl = "CatTreeList.aspx?DatabaseId=" & DatabaseId & "&CategoryId=&ActorType=" & ActorType & "&t=" & Request("t") & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                    Else
                        PreviousText.Visible = False
                        PreviousImg.Visible = False
                    End If
                    If PageNumberDDL.SelectedValue < pageCount Then
                        NextText.Visible = True
                        NextImg.Visible = True
                        NextLink.NavigateUrl = "CatTreeList.aspx?DatabaseId=" & DatabaseId & "&CategoryId=&ActorType=" & ActorType & "&t=" & Request("t") & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                    Else
                        NextText.Visible = False
                        NextImg.Visible = False
                    End If

                    Dim i As Integer = 0
                    link = ""
                    sb.Append("<ul class=""list"">")

                    Dim CatArr() As String

                    For i = 0 To PageSize - 1

                        CatArr = GetCategoryIdandName(myDataRow("REPORT_ID"), myConnection)

                        sb.Append("<li>")

                        sb.Append("<a href=""CatTreeContent.aspx?ReportId=" & myDataRow("REPORT_ID") & "&DatabaseId=" & DatabaseId)

                        sb.Append("&CategoryId=" & CatArr(0))

                        sb.Append("&ActorType=002"">" & myDataRow("SUBJECT") & "</a>")

                        sb.Append("<span>[" & CatArr(1) & "]</span>[" & myDataRow("ONLINE_DATE") & "]")

                        If Request("t") = "b" Then
                            sb.Append("[" & myDataRow("CLICK_COUNT") & "]")
                        End If

                        sb.Append("</li>")
                        position += 1
                        If myTable.Rows.Count <= position Then
                            Exit For
                        End If
                        myDataRow = myTable.Rows.Item(position)
                    Next
                    sb.Append("</ul>")
                    sb.Append("</div>")

                    TableText.Text = sb.ToString()

                End If


            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try

        Else
            PageNumberDDL.SelectedIndex = 0
            PageSizeDDL.SelectedIndex = 0
            TableText.Text = "<div class=""nodata"">本層節點目前無資料，您可點選上方子目錄資料項目、或點選左欄其他知識庫節點。</div>"
        End If

    End Sub

    Function GetCategoryIdandName(ByVal id As String, ByRef conn As SqlConnection) As Array

        Dim arr() As String = {"", ""}
        Dim reader As SqlDataReader
        Dim command As SqlCommand
        Dim sql As String = ""
        Try
            sql = "SELECT DISTINCT CATEGORY.CATEGORY_ID, CATEGORY.CATEGORY_NAME FROM REPORT INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN CATEGORY "
            sql &= "ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID WHERE (REPORT.REPORT_ID = @id) AND (CAT2RPT.DATA_BASE_ID = 'DB020') "
            command = New SqlCommand(sql, myConnection)
            command.Parameters.AddWithValue("@id", id)
            reader = command.ExecuteReader()

            If reader.Read() Then
                arr(0) = reader("CATEGORY_ID")
                arr(1) = reader("CATEGORY_NAME")
            End If
            If Not reader.IsClosed Then
                reader.Close()
            End If
        Catch ex As Exception
            Response.Write(ex.ToString)
        End Try

        Return arr

    End Function

End Class
