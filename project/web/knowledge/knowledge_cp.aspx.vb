Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_cp
    Inherits System.Web.UI.Page

    Protected innerTag As String = ""
    Protected accountTag As String = ""
    Dim ArticleId As String = ""
    Dim ArticleType As String = ""
    Dim CategoryId As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim MemberId As String = ""
        Dim MemberActor As String = ""
        Dim gardening As String = WebUtility.GetStringParameter("gardening", String.Empty)

        If Request.QueryString("ArticleType") <> "A" And Request.QueryString("ArticleType") <> "B" And Request.QueryString("ArticleType") <> "C" _
            And Request.QueryString("ArticleType") <> "D" And Request.QueryString("ArticleType") <> "E" Then
            'Response.Redirect("/knowledge/knowledge.aspx")
            'Response.End()
            ArticleType = "A"
        Else
            ArticleType = Request.QueryString("ArticleType")
        End If

        If Request.QueryString("CategoryId") <> "A" And Request.QueryString("CategoryId") <> "B" And Request.QueryString("CategoryId") <> "C" _
            And Request.QueryString("CategoryId") <> "D" And Request.QueryString("CategoryId") <> "E" And Request.QueryString("CategoryId") <> "F" Then
            'Response.Redirect("/knowledge/knowledge.aspx")
            'Response.End()
            CategoryId = ""
        Else
            CategoryId = Request.QueryString("CategoryId")
        End If

        If Request.QueryString("ArticleId") = "" Then
            Response.Redirect("/knowledge/knowledge.aspx")
            Response.End()
        Else
            ArticleId = HttpUtility.UrlEncode(Request.QueryString("ArticleId"))
        End If

        'Response.Write (ArticleId)
        'added by Joey,問題單回應 問題、討論的區分
        Me.ARTICLE_ID_Question.value = ArticleId
        '-----------------------------------------------------------------------------
        If Session("memID") = "" Or Session("memID") Is Nothing Then
            MemberId = ""
            MemberActor = ""
        Else
            MemberId = Session("memID").ToString()
            MemberActor = Session("gstyle").ToString()
        End If
        '-----------------------------------------------------------------------------

        '---kpi use---reflash wont add grade---
        If Request.QueryString("kpi") <> "0" Then
            '---start of kpi user---20080911---vincent---
            If Session("memID") IsNot Nothing Or Session("memID") <> "" Then
                Dim browse As New KPIBrowse(Session("memID"), "browseQuestionCP", ArticleId)
                browse.HandleBrowse()
            End If
            '---end of kpi use---
            Dim relink As String = ""
            If gardening <> "" Then
                relink = "/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&kpi=0&gardening=true"
            Else
                relink = "/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&kpi=0"
            End If
            Response.Redirect(relink)
            Response.End()
        End If
        '---end of kpi use---

        PathText.Text = "&gt;<a href=""/knowledge/knowledge_lp.aspx"">知識分類</a>"

        If CategoryId = "A" Then
            PathText.Text &= "&gt;<a href=""/knowledge/knowledge_lp.aspx?ArticleType=A&CategoryId=A"">農</a>"
        ElseIf CategoryId = "B" Then
            PathText.Text &= "&gt;<a href=""/knowledge/knowledge_lp.aspx?ArticleType=A&CategoryId=B"">林</a>"
        ElseIf CategoryId = "C" Then
            PathText.Text &= "&gt;<a href=""/knowledge/knowledge_lp.aspx?ArticleType=A&CategoryId=C"">漁</a>"
        ElseIf CategoryId = "D" Then
            PathText.Text &= "&gt;<a href=""/knowledge/knowledge_lp.aspx?ArticleType=A&CategoryId=D"">牧</a>"
        ElseIf CategoryId = "E" Then
            PathText.Text &= "&gt;<a href=""/knowledge/knowledge_lp.aspx?ArticleType=A&CategoryId=E"">其他</a>"
        ElseIf CategoryId = "F" Then
            PathText.Text &= "&gt;<a href=""/knowledge/knowledge_lp.aspx?ArticleType=A&CategoryId=F"">產銷經營管理系統</a>"
        End If

        TraceAddText.Text = "<a href=""/knowledge/knowledge_trace_add.aspx?ArticleId=" & ArticleId & """ class=""Track"">放入追蹤</a>"

        If gardening <> "" Then
            BackText.Text = "<a href=""javascript:history.back()"" class=""Back"">回上一頁</a>"
        Else
            BackText.Text = "<a href=""/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """ class=""Back"">回上一頁</a>"
        End If

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim sb As New StringBuilder
        Dim HavePros As String = ""
        Dim myTable As DataTable
        Dim myDataRow As DataRow
        Dim tagarray As Array
        Dim DiscussFlag As Boolean = True

        Dim Activity As New ActivityFilter(WebConfigurationManager.AppSettings("ActivityId"), "")
        Dim ActivityFlag As Boolean = Activity.CheckActivity

        myConnection = New SqlConnection(ConnString)

        '計算討論則數 
        Dim totalDiscuss As Integer = 0
        Try
            sqlString = "SELECT COUNT(*) FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
            sqlString &= "WHERE (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) "
            sqlString &= "AND (KnowledgeForum.Status = 'N') "
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))

            totalDiscuss = myCommand.ExecuteScalar

            myCommand.Dispose()

        Catch ex As Exception
            If Request.QueryString("debug") = "true" Then
                Response.Write(ex.ToString)
                Response.End()
            End If
            Response.Redirect("/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
            Response.End()
        Finally
            myConnection.Close()
        End Try
        '計算討論則數 end

        Try
            myConnection.Open()

            '---更新瀏覽數---
            '把count移到viewcounter.aspx 
            'sqlString = "UPDATE KnowledgeForum SET BrowseCount = BrowseCount + 1 WHERE gicuitem = @agicuitem"
            'myCommand = New SqlCommand(sqlString, myConnection)
            'myCommand.Parameters.AddWithValue("@agicuitem", ArticleId)
            'myCommand.ExecuteNonQuery()
            'myCommand.Dispose()

            'sqlString = "UPDATE CuDTGeneric SET ClickCount = (SELECT BrowseCount FROM KnowledgeForum WHERE gicuitem = @agicuitem) WHERE icuitem = @agicuitem"
            'myCommand = New SqlCommand(sqlString, myConnection)
            'myCommand.Parameters.AddWithValue("@agicuitem", ArticleId)
            'myCommand.ExecuteNonQuery()
            'myCommand.Dispose()

            sqlString = "SELECT KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount, KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, "
            sqlString &= "KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount,  CuDTGeneric.xNewWindow, ISNULL(CuDTGeneric.vGroup, '') AS vGroup, "
            sqlString &= "KnowledgeForum.HavePros, KnowledgeForum.ParentIcuitem, CuDTGeneric.iCUItem, CuDTGeneric.iCTUnit, CuDTGeneric.sTitle, CuDTGeneric.topCat, "
            sqlString &= "CuDTGeneric.xKeyword, CuDTGeneric.xPostDate, ISNULL(CuDTGeneric.xBody, '') AS xBody, isnull(Member.realname,'') as realname, ISNULL(Member.nickname, '') AS nickname FROM KnowledgeForum INNER JOIN CuDTGeneric ON "
            sqlString &= "KnowledgeForum.gicuitem = CuDTGeneric.iCUItem left JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCUItem = @ArticleId) "
            sqlString &= "AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N')"

            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@ArticleId", ArticleId)
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
            myReader = myCommand.ExecuteReader()

            If myReader.Read() Then

                If Session("BrowseTitleNew") <> myReader("sTitle") Then
                    Session("BrowseTitleNew") = myReader("sTitle")
                    Session("BrowseLinkNew") = "/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId
                End If

                If myReader("xNewWindow") = "Y" Then
                    DiscussFlag = False
                End If

                If myReader("xPostDate") IsNot DBNull.Value Then
                    xPostDateText.Text = Date.Parse(myReader("xPostDate")).ToShortDateString()
                End If

                sb.Append("<h3>標題：" & myReader("sTitle") & "</h3>")
                sb.Append("<table class=""subtable"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""排版用表格"">")
                sb.Append("<tr><td class=""subleft"">")

                sb.Append("<div class=""problem"">" & ReplaceAndFindKeyword(myReader("xBody")) & "&nbsp;</div>")

                sb.Append(GenerateQuestionAdditional(myReader("iCUItem")))

                sb.Append(GetPicInfo(myReader("iCUItem"), 300))


                sb.Append("<td class=""subright""><div><ul>")

                If ActivityFlag Then
                    If myReader("vGroup") = "A" And myReader("xNewWindow") <> "Y" Then
                        sb.Append("<li>本問題為「知識問答你˙我˙他」活動問題")
                        sb.Append("<div class=""staricon"">")
                        sb.Append("<img class=rating src=""../xslGip/style3/images/good.gif"" align=top>")
                        sb.Append("</div></li><br/><br/>")
                        '回上一頁修改
                        BackText.Text = "<a href=""/knowledge/knowledge_alp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """ class=""Back"">回上一頁</a>"
                    End If
                End If

                sb.Append("<li>發問者：")
                If myReader("nickname") <> "" Then
                    sb.Append(myReader("nickname").ToString.Trim)
                Else
                    Dim name As String = myReader("realname").ToString.Trim
                    If name.Length > 0 Then
                        If instr(name, "&#") > 0 And instr(name, ";") > 1 Then
                            name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
                        ElseIf instr(name, "&#") > 0 And instr(name, ";") < 1 Then
                            name = name.Substring(0, 7) & "＊" & name.Substring(name.Length - 7, 7)
                        Else
                            name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
                        End If
                    End If
                    sb.Append(name)
                End If
                sb.Append("</li>")
                sb.Append("<li>Tags：")
                If myReader("xKeyword") IsNot DBNull.Value Then
                    tagarray = myReader("xKeyword").ToString().Split(";")
                    If tagarray.Length > 0 Then
                        For Each tag As String In tagarray
                            If tag <> "" Then
                                sb.Append(tag.ToString() & "、")
                            End If
                        Next
                    Else
                        sb.Append("無")
                    End If
                Else
                    sb.Append("無")
                End If
                sb.Append("</li>")

                'sb.Append("<li>瀏覽數：" & myReader("BrowseCount") & "</li><li>討論數：" & myReader("DiscussCount") & "</li>")
                sb.Append("<li>瀏覽數：" & myReader("BrowseCount") & "</li><li>討論數：" & totalDiscuss.ToString & "</li>")
                sb.Append("<li>追蹤數：" & myReader("TraceCount") & "</li><li>意見數：" & myReader("CommandCount") & "</li>")
                sb.Append(ConcatStarString("0", Integer.Parse(myReader("GradeCount")), Integer.Parse(myReader("GradePersonCount"))))
                sb.Append("</ul></td></tr></table>")

                QuestionText.Text = sb.ToString()

                '---若有專家補充---
                HavePros = myReader("HavePros")

                sb = Nothing

            Else
                Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "linkerror", Script.Replace("{0}", "此文章已被刪除!!"))
                Exit Sub
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

        Catch ex As Exception
            If Request.QueryString("debug") = "true" Then
                Response.Write(ex.ToString)
                Response.End()
            End If
            Response.Redirect("/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
            Response.End()
        Finally
            myConnection.Close()
        End Try

        If HavePros = "Y" Then
            sb = New StringBuilder
            Try
                sqlString = "SELECT Member.account, Member.realname, ISNULL(Member.nickname, '') AS nickname, ISNULL(Member.photo, '') AS photo, CuDTGeneric.xBody, CuDTGeneric.xPostDate, CuDTGeneric.iCUItem "
                sqlString &= "FROM CuDTGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account INNER JOIN KnowledgeForum ON KnowledgeForum.gicuitem = CuDTGeneric.iCUItem "
                sqlString &= "WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N') AND (KnowledgeForum.ParentIcuitem = @icuitem) "
                sqlString &= "AND (CuDTGeneric.iCTUnit = @ictunit) ORDER BY CuDTGeneric.xPostDate DESC "
                'Response.Write(sqlString)
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
                myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
                myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("ProfessionAnswerCtUnitId"))
                myReader = myCommand.ExecuteReader()
                While myReader.Read()

                    sb.Append("<div class=""experts""><h4>駐站專家補充</h4>")
                    sb.Append("<div class=""exphoto""><div class=""phoimg"">")
                    If myReader("photo") = "" Then
                        sb.Append("<img src=""images/expert.gif"" alt=""圖片說明"" />")
                    Else
                        'sb.Append("<img src=""/Public/Data/" & myReader("photo") & """ alt=""圖片說明"" />")
                    End If
                    sb.Append("</div>")
                    If myReader("nickname") <> "" Then
                        sb.Append("<p>" & myReader("nickname").ToString.Trim)
                    Else
                        Dim name As String = myReader("realname").ToString.Trim
                        If instr(name, "&#") > 0 And instr(name, ";") > 1 Then
                            name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
                        ElseIf instr(name, "&#") > 0 And instr(name, ";") < 1 Then
                            name = name.Substring(0, 7) & "＊" & name.Substring(name.Length - 7, 7)
                        Else
                            name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
                        End If
                        sb.Append("<p>" & name)
                    End If
                    sb.Append(" 發表於")
                    If myReader("xPostDate") IsNot DBNull.Value Then
                        sb.Append(Date.Parse(myReader("xPostDate")).ToShortDateString())
                    End If
                    sb.Append("</p>")
                    sb.Append("<div class=""float""></div></div>")
                    'sb.Append("<p>專長：<a href=""#"">蝴蝶蘭</a>、<a href=""#"">教學</a>、<a href=""#"">音樂</a></p>")
                    'sb.Append("<div class=""float""><a href=""#"">更多</a></div></div>")
                    'If myReader("xBody").ToString().Length > 200 Then
                    'sb.Append("<p>" & myReader("xBody").ToString().Substring(0, 200).Replace(Chr(13), "<br />"))
                    'Else
                    sb.Append("<div style='float:left'>")
                    Dim imagesHtml As String = ""
                    sb.Append("<p>" & ReplaceAndFindKeyword(myReader("xBody").replace(Chr(13), "<br />")))
                    sb.Append("</p>")
                    imagesHtml = GetPicInfo(myReader("iCUItem"), 380)
                    sb.Append(imagesHtml)
                    If myReader("account").ToString().Trim() = MemberId.Trim() Then
                        sb.Append("<div><input id=btn_" & myReader("iCUItem") & " type=button value='刪除回覆' onclick=btnDel(this) /> ")
                        If imagesHtml <> "" Then
                            sb.Append("&nbsp;&nbsp;<input id=btnpic_" & myReader("iCUItem") & " type=button value='刪除圖片' onclick=btnDelPic(this) /> ")
                        End If
                        sb.Append("</div>")
                    End If
                    'sb.Append(GetPicInfo(myReader("iCUItem"), 520))
                    'End If
                    'sb.Append("...<a href=""/knowledge/knowledge_professional.aspx?ArticleId=" & ArticleId & "&DArticleId=" & myReader("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>詳全文</a></p></div>")

                    sb.Append("</div></div>")

                End While
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

                ProsText.Text = sb.ToString()
                sb = Nothing

            Catch ex As Exception
                If Request.QueryString("debug") = "true" Then
                    Response.Write(ex.ToString)
                    Response.End()
                End If
                Response.Redirect("/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
                Response.End()
            Finally
                myConnection.Close()
            End Try
        Else
            ProsText.Text = ""
        End If




        Try
            sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, KnowledgeForum.DiscussCount, "
            sqlString &= "KnowledgeForum.CommandCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, Member.realname, ISNULL(Member.nickname, '') AS nickname, Member.account "
            sqlString &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
            sqlString &= "WHERE (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) "
            sqlString &= "AND (KnowledgeForum.Status = 'N') ORDER BY CuDTGeneric.xPostDate DESC"

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
            myReader = myCommand.ExecuteReader()

            myTable = New DataTable
            myTable.Load(myReader)

            Dim Total As Integer = myTable.Rows.Count()
            Dim i As Integer = 0
            sb = New StringBuilder

            sb.Append("<div class=""Dis""><h4>討論(" & totalDiscuss.ToString & ")</h4>")
            'sb.Append("<div class=""slidertable""><table border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""排版用表格"">")
            'sb.Append("<tr><td>摘要字數：(少)</td><td valign=""middle"" class=""slider""><img src=""images/btn_slider.gif"" alt=""sliderbar"" /></td>")
            'sb.Append("<td>(多)</td></tr></table></div>")
            '功能連結

            sb.Append("<div class=""float2"">")
            If MemberId <> "" And MemberId IsNot Nothing And MemberActor = "005" Then
                sb.Append("<a href=""/knowledge/knowledge_professional_add.aspx?ArticleId=" & Request.QueryString("ArticleId") & _
                        "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>專家補充</a>")

            End If
            If DiscussFlag Then
                sb.Append("<a href=""/knowledge/knowledge_discuss.aspx?ArticleId=" & Request.QueryString("ArticleId") & _
                            "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>參加討論</a>")
            End If
            sb.Append("</div>")

            If Total <> 0 Then

                For i = 0 To Total
                    myDataRow = myTable.Rows.Item(i)
                    sb.Append("<div class=""DisList""><ul class=""Issued"" >")

                    Dim showName As String
                    If myDataRow("nickname") <> "" Then
                        showName = myDataRow("nickname").ToString.Trim
                        'sb.Append("<li>" & myDataRow("nickname").ToString.Trim & "</li>")
                    Else
                        Dim name As String = myDataRow("realname").ToString.Trim
                        If instr(name, "&#") > 0 And instr(name, ";") > 1 Then
                            name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
                        ElseIf instr(name, "&#") > 0 And instr(name, ";") < 1 Then
                            name = name.Substring(0, 7) & "＊" & name.Substring(name.Length - 7, 7)
                        Else
                            name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
                        End If
                        showName = name
                        'sb.Append("<li>" & name & "</li>")
                    End If
                    'sb.Append("<li><a >" & showName & "</a></li>")	  

                    '===========在知識家的討論串中, 顯示個人等級 2009.08.18 by ivy
                    Dim account As String = myDataRow("account").ToString.Trim
                    Dim sqlMemLevel As String = ""
                    Dim gradeLevel As String = "", gradeDesc As String = ""
                    Dim memLevelCmd As SqlCommand
                    Dim memLevelReader As SqlDataReader
                    sqlMemLevel = "SELECT top 1 mCode FROM CodeMain WHERE (codeMetaID = 'gradelevel') "
                    sqlMemLevel &= "AND ((SELECT calculateTotal FROM MemberGradeSummary WHERE memberId = @account)>= mValue) ORDER BY mSortValue DESC"
                    memLevelCmd = New SqlCommand(sqlMemLevel, myConnection)
                    memLevelCmd.Parameters.AddWithValue("@account", account)
                    gradeLevel = memLevelCmd.ExecuteScalar()

                    If gradeLevel = "1" Then
                        gradeDesc = "入門級會員"
                    ElseIf gradeLevel = "2" Then
                        gradeDesc = "進階級會員"
                    ElseIf gradeLevel = "3" Then
                        gradeDesc = "高手級會員"
                    ElseIf gradeLevel = "4" Then
                        gradeDesc = "達人級會員"
                    Else
                        gradeDesc = "入門級會員"
                    End If



                    sb.Append("<input type='hidden' id='" + "d_account_" + myDataRow("iCUItem").ToString() + "' value='" + myDataRow("account").ToString.Trim + "'></input>")

                    If String.IsNullOrEmpty(myDataRow("nickname").ToString.Trim) Then
                        sb.Append("<input type='hidden' id='" + "d_nickname_" + myDataRow("iCUItem").ToString() + "' value='" + myDataRow("realname").ToString.Trim + "'></input>")
                    Else
                        sb.Append("<input type='hidden' id='" + "d_nickname_" + myDataRow("iCUItem").ToString() + "' value='" + myDataRow("nickname").ToString.Trim + "'></input>")
                    End If



                    'JIRA COAKM-36 討論中可以看到參予討論人員(所有會員，不限制是達人)的平均評價
                    sb.Append("<li>" & MemberInfo(myDataRow("account").ToString.Trim, showName, Convert.ToInt32(gradeLevel)) & "</li>")
                    sb.Append("(" & gradeDesc & ")")

                    sb.Append("<li>發表於 ")
                    If myDataRow("xPostDate") IsNot DBNull.Value Then
                        sb.Append(Date.Parse(myDataRow("xPostDate")).ToShortDateString())
                    End If
                    sb.Append("</li>")

                    sb.Append("<li>意見數：" & myDataRow("CommandCount") & "</li>")
                    sb.Append(ConcatStarString("1", Integer.Parse(myDataRow("GradeCount")), Integer.Parse(myDataRow("GradePersonCount"))))
                    'added by Joey, 2009/09/21, 討論區新增問題反應連結

                    sb.Append("<li><a href=""#"" onclick=""Discuss(" + myDataRow("iCUItem").ToString() + ")"">檢舉</a></li>")
                    sb.Append("<input type=""hidden""  name=""type"" value=""2"">")
                    sb.Append("<li><input type=""hidden""  name=""ARTICLE_ID"" value=" & myDataRow("iCUItem") & "></li>")


                    sb.Append("</ul><p>")
                    'If myDataRow("xBody").ToString().Length > 200 Then
                    'sb.Append(myDataRow("xBody").ToString().Substring(0, 200).Replace("<", "&lt;").Replace(">", "&gt;").Replace(Chr(13), "<br />"))
                    ' Else
                    sb.Append(ReplaceAndFindKeyword(myDataRow("xBody").replace(Chr(13), "<br />")))
                    sb.Append(GetPicInfo(myDataRow("iCUItem"), 480))
                    ' End If
                    sb.Append("&nbsp;<a href=""/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & myDataRow("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&type=1"">發表意見</a>")
                    '意見數大於0 才有此選項
                    If myDataRow("CommandCount") > 0 Then
                        sb.Append("&nbsp;<a href=""/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & myDataRow("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&type=2"">瀏覽意見</a>")
                    End If
                    'sb.Append("...<a href=""/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & myDataRow("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>詳全文</a></p></div>")
                    sb.Append("</p></div>")

                    If myTable.Rows.Count = i + 1 Then
                        Exit For
                    End If
                Next
                'sb.Append("<div class=""float""><a href=""/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>看全部討論</a></div></div>")
            End If

            DiscussText.Text = sb.ToString
            sb = Nothing

            myCommand.Dispose()

        Catch ex As Exception
            If Request.QueryString("debug") = "true" Then
                Response.Write(ex.ToString)
                Response.End()
            End If
            Response.Redirect("/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
            Response.End()
        Finally
            myConnection.Close()
        End Try

        '相似問題
        SimilarProblemText.Text = GenerateSimilarProblem(ArticleId)

    End Sub

    Function GenerateQuestionAdditional(ByVal icuitem As String)

        Dim str As String = ""
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader

        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT CuDTGeneric.xPostDate, CuDTGeneric.xBody FROM KnowledgeForum INNER JOIN CuDTGeneric ON KnowledgeForum.gicuitem = CuDTGeneric.iCUItem "
            sqlString &= "WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (CuDTGeneric.iCTUnit = @ictunit) AND (KnowledgeForum.ParentIcuitem = @icuitem) "
            sqlString &= "AND (KnowledgeForum.Status = 'N')"
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", icuitem)
            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeAdditionalCtUnitId"))
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
            myReader = myCommand.ExecuteReader()

            While myReader.Read
                str &= "<div class=""adddate"">補充說明發表日期："
                str &= Date.Parse(myReader("xPostDate")).ToShortDateString
                str &= "</div>"
                str &= "<div class=""problem2"">" & ReplaceAndFindKeyword(myReader("xBody").replace(Chr(13), "<br />")) & "</div>"
            End While

            If Not myReader.IsClosed Then
                myReader.Close()
            End If

            myCommand.Dispose()

        Catch ex As Exception
            Response.Write(ex.ToString())
            Response.End()
        Finally
            myConnection.Close()
        End Try

        Return str

    End Function

    Function GetPicInfo(ByVal aid As String, ByVal maxWidth As Integer) As String

        Dim str As String = ""
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader

        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT * FROM KnowledgePicInfo INNER JOIN CuDTGeneric ON KnowledgePicInfo.parentIcuitem = CuDTGeneric.iCUItem "
            sqlString &= "WHERE (CuDTGeneric.iCUItem = @icuitem) AND (KnowledgePicInfo.picStatus = 'Y')"
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", aid)
            myReader = myCommand.ExecuteReader()

            While myReader.Read
                str &= "<div style='width:" & maxWidth & "px'>"
                str &= "<a href=""" & myReader("picPath") & """ alt=""" & myReader("picTitle") & """ target='_blank'><img src=""" & myReader("picPath") & """ alt=""" & myReader("picTitle") & """ style='max-width:" & maxWidth & "px' /></a>"
                str &= "<p style='color:#666666'>" & myReader("picTitle") & "</p>"
                str &= "</div>"
            End While

            If Not myReader.IsClosed Then
                myReader.Close()
            End If

            myCommand.Dispose()

        Catch ex As Exception
            Response.Write(ex.ToString())
            Response.End()
        Finally
            myConnection.Close()
        End Try

        Return str

    End Function

    Function ConcatStarString(ByVal type As String, ByVal GradeCount As Integer, ByVal PersonCount As Integer) As String

        Dim dresult As Double = 0

        If PersonCount = 0 Then
            dresult = GradeCount / 1
        Else
            dresult = GradeCount / PersonCount
        End If

        Dim str As String = "<li>平均評價：<div class=""staricon"">"
        If dresult = 0 Then
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 0 And dresult < 0.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 0.5 And dresult <= 1 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 1 And dresult < 1.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 1.5 And dresult <= 2 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 2 And dresult < 2.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 2.5 And dresult <= 3 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 3 And dresult < 3.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 3.5 And dresult <= 4 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 4 And dresult < 4.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
        ElseIf dresult >= 4.5 And dresult <= 5.0 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
        End If

        str &= "</div>"
        If type = "0" Then
            str &= "<div class=""float"">(" & PersonCount & "人評價)</div></li>"
        ElseIf type = "1" Then
            str &= "(" & PersonCount & "人評價)</li>"
        End If

        Return str

    End Function

    Function ReplaceAndFindKeyword(ByVal strTemp As String) As String
        strTemp = strTemp.Replace(Chr(13), "<br />")
        Dim string1
        string1 = PediaUtility.ReplacePedia(strTemp)
        strTemp = string1(0)
        innerTag &= string1(1)

        Return strTemp
    End Function

    Function MemberInfo(ByVal account As String, ByVal name As String, ByVal gradeLevel As Integer) As String

        Dim string1
        string1 = MemberInfoUtility.ReplaceMemberName(account, name, gradeLevel)
        account = string1(0)
        accountTag &= string1(1)

        Return account

    End Function


    '相似問題
    Function GenerateSimilarProblem(ByVal icuItem As String) As String
        Dim str As String = ""
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim index As Integer = 1

        myConnection = New SqlConnection(ConnString)
        Try
            Dim sqlStr As StringBuilder = New StringBuilder()
            sqlStr.Append(" SELECT Top (5) iCUItem ,topCat ,sTitle ,vGroup,xPostDate,siteId,B.id_count,C.mValue  " & VbCrLf)
            sqlStr.Append(" FROM  " & VbCrLf)
            sqlStr.Append(" (  " & VbCrLf)
            sqlStr.Append(" 	SELECT iCUItem,topCat,sTitle,vGroup,xPostDate,siteId  " & VbCrLf)
            sqlStr.Append(" 	FROM CuDTGeneric  " & VbCrLf)
            sqlStr.Append(" 	INNER JOIN 		KnowledgeForum  " & VbCrLf)
            sqlStr.Append(" 	ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem and (KnowledgeForum.status = 'N')  " & VbCrLf)
            sqlStr.Append(" 	Where CuDTGeneric.fCTUPublic='Y' and siteId= '3' " & VbCrLf)
            sqlStr.Append(" ) A  " & VbCrLf)
            sqlStr.Append(" LEFT JOIN  " & VbCrLf)
            sqlStr.Append(" (  " & VbCrLf)
            sqlStr.Append(" 	select REPORT_ID,count(REPORT_ID) As id_count  " & VbCrLf)
            sqlStr.Append(" 	FROM coa.dbo.REPORT_KEYWORD_FREQUENCY  " & VbCrLf)
            sqlStr.Append(" 	where REPORT_ID != @icuitem and Keyword in (  " & VbCrLf)
            sqlStr.Append(" 		SELECT Keyword  " & VbCrLf)
            sqlStr.Append(" 		FROM coa.dbo.REPORT_KEYWORD_FREQUENCY  " & VbCrLf)
            sqlStr.Append(" 		where REPORT_ID = @icuitem " & VbCrLf)
            sqlStr.Append(" 		AND NOT KEYWORD IN ( " & VbCrLf)
            sqlStr.Append(" 			select top 1 KEYWORD " & VbCrLf)
            sqlStr.Append(" 			from coa.dbo.REPORT_KEYWORD_FREQUENCY " & VbCrLf)
            sqlStr.Append(" 			GROUP BY KEYWORD " & VbCrLf)
            sqlStr.Append(" 			ORDER BY COUNT(*) DESC	 " & VbCrLf)
            sqlStr.Append(" 		) 	AND  exists(SELECT STITLE FROM CUDTGENERIC WHERE ICUITEM=@icuitem and STITLE  LIKE '%' + KEYWORD + '%' and xbody  LIKE '%' + KEYWORD + '%' ) " & VbCrLf)
            sqlStr.Append(" 	) AND  exists(SELECT STITLE FROM CUDTGENERIC WHERE ICUITEM=coa.dbo.REPORT_KEYWORD_FREQUENCY.REPORT_ID and STITLE  LIKE '%' + KEYWORD + '%' and xbody  LIKE '%' + KEYWORD + '%' ) " & VbCrLf)
            sqlStr.Append(" 	group by REPORT_ID  " & VbCrLf)
            sqlStr.Append(" )B  " & VbCrLf)
            sqlStr.Append(" ON A.iCUITem = B.REPORT_ID  " & VbCrLf)
            sqlStr.Append(" LEFT JOIN (  " & VbCrLf)
            sqlStr.Append(" 	select mValue,mCode  " & VbCrLf)
            sqlStr.Append(" 	from CodeMain  " & VbCrLf)
            sqlStr.Append(" 	where codeMetaID = 'KnowledgeType'  " & VbCrLf)
            sqlStr.Append(" )C  " & VbCrLf)
            sqlStr.Append(" on A.topCat = C.mCode  " & VbCrLf)
            sqlStr.Append(" where NOT B.REPORT_ID IS NULL " & VbCrLf)
            sqlStr.Append(" order by id_count desc , xPostDate desc " & VbCrLf)
            myConnection.Open()
            myCommand = New SqlCommand(sqlStr.ToString(), myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", icuItem)
            myReader = myCommand.ExecuteReader()
            While myReader.Read
                If index = 1 Then
                    str &= "<table class=""similar"">"
                    str &= "<tr><th colSpan=""4"">相似問題</th></tr>"
                End If
                str &= "<tr><td width=""5%"">" & index & "</td>"
                str &= "<td width=""74%""><a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("iCUItem") & "&ArticleType=A&CategoryId=" & myReader("topCat") & """>" & myReader("sTitle") & "</a></td>"
                str &= "<td width=""6%"" ALIGN=""CENTER"">" & myReader("mValue") & "</td><td width=""15%"">[" & Date.Parse(myReader("xPostDate")).ToShortDateString & "]</td></tr>"
                index = index + 1
            End While
            If str <> "" Then
                str &= "</table>"
            End If
        Catch ex As Exception
            Response.Write(ex.ToString())
            Response.End()
        Finally
            myConnection.Close()
        End Try
        Return str
    End Function

    Protected Sub btnDelProsText_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDelProsText.Click
        Response.Write(hidProsText.Value)
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader

        myConnection = New SqlConnection(ConnString)
        Try
            myConnection.Open()

            sqlString = "select * from KnowledgePicInfo where ParentIcuitem = @parentIcuitem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@parentIcuitem", Convert.ToInt32(hidProsText.Value))
            myReader = myCommand.ExecuteReader()
            While myReader.Read
                Dim picFile As String = myReader("picPath").ToString()
                Dim realPicFile As String = Server.MapPath("/public/Data/") + picFile.Substring(picFile.LastIndexOf("/") + 1)
                Response.Write(Server.MapPath("/public/Data/") + picFile.Substring(picFile.LastIndexOf("/") + 1))
                System.IO.File.Delete(realPicFile)
            End While
            If Not myReader.IsClosed Then
                myReader.Close()
            End If

            sqlString = "delete from KnowledgePicInfo where ParentIcuitem = @parentIcuitem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@parentIcuitem", Convert.ToInt32(hidProsText.Value))
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()

            sqlString = "delete from KnowledgeForum where ParentIcuitem = @parentIcuitem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@parentIcuitem", Convert.ToInt32(hidProsText.Value))
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()

            sqlString = "UPDATE HyftdIndexDelete SET isDelete = 1 WHERE iCUItem = @iCUItem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@iCUItem", Convert.ToInt32(hidProsText.Value))
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()

            sqlString = "delete from CuDTGeneric where iCUItem = @iCUItem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@iCUItem", Convert.ToInt32(hidProsText.Value))
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()

            '2010/12/29 檢查是否還有專家補充 沒有則修改欄位
            sqlString = " IF NOT EXISTS( "
            sqlString &= " select * from  CuDTGeneric where iCUItem in( "
            sqlString &= " select gicuitem from KnowledgeForum where ParentIcuitem = @ArticleId ) and ictunit = @ictunit  "
            sqlString &= " ) "
            sqlString &= " begin "
            sqlString &= " 	Update KnowledgeForum set HavePros = 'N' where giCUItem = @ArticleId "
            sqlString &= " end  "
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@ArticleId", ArticleId)
            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("ProfessionAnswerCtUnitId"))
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()


        Catch ex As Exception
            Response.Write(ex.ToString())
            Response.End()
        Finally
            myConnection.Close()
            Response.Redirect("/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&kpi=0")
            Response.End()
        End Try
    End Sub

    Protected Sub btnDelPic_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDelPic.Click
        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader

        myConnection = New SqlConnection(ConnString)
        Try
            myConnection.Open()

            sqlString = "select * from KnowledgePicInfo where ParentIcuitem = @parentIcuitem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@parentIcuitem", Convert.ToInt32(hidProsText.Value))
            myReader = myCommand.ExecuteReader()
            While myReader.Read
                Dim picFile As String = myReader("picPath").ToString()
                Dim realPicFile As String = Server.MapPath("/public/Data/") + picFile.Substring(picFile.LastIndexOf("/") + 1)
                Response.Write(Server.MapPath("/public/Data/") + picFile.Substring(picFile.LastIndexOf("/") + 1))
                System.IO.File.Delete(realPicFile)
            End While
            If Not myReader.IsClosed Then
                myReader.Close()
            End If

            sqlString = "delete from KnowledgePicInfo where ParentIcuitem = @parentIcuitem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@parentIcuitem", Convert.ToInt32(hidProsText.Value))
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()

        Catch ex As Exception
            Response.Write(ex.ToString())
            Response.End()
        Finally
            myConnection.Close()
            Response.Redirect("/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&kpi=0")
            Response.End()
        End Try
    End Sub
End Class
