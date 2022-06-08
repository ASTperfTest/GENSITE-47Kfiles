Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class myknowledge_trace_lp
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ArticleType As String
        Dim CategoryId As String
        Dim Sort As String
        Dim Keyword As String
        Dim sb As StringBuilder
        Dim MemberId As String
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>"

        '-----------------------------------------------------------------------------
        If Session("memID") = "" Or Session("memID") Is Nothing Then
            MemberId = ""
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
            Exit Sub
        Else
            MemberId = Session("memID").ToString()
        End If
        '-----------------------------------------------------------------------------

        If Request.QueryString("ArticleType") <> "A" And Request.QueryString("ArticleType") <> "B" And Request.QueryString("ArticleType") <> "C" _
                    And Request.QueryString("ArticleType") <> "D" And Request.QueryString("ArticleType") <> "E" And Request.QueryString("ArticleType") <> "F" Then
            ArticleType = "A"
        Else
            ArticleType = Request.QueryString("ArticleType")
        End If

        If Request.QueryString("CategoryId") <> "A" And Request.QueryString("CategoryId") <> "B" And Request.QueryString("CategoryId") <> "C" _
                    And Request.QueryString("CategoryId") <> "D" And Request.QueryString("CategoryId") <> "E" And Request.QueryString("CategoryId") <> "F" Then
            CategoryId = ""
        Else
            CategoryId = Request.QueryString("CategoryId")
        End If

        If Request.QueryString("Sort") <> "A" And Request.QueryString("Sort") <> "D" Then
            Sort = "D"
        Else
            Sort = Request.QueryString("Sort")
        End If

        If Request.QueryString("keyword") IsNot Nothing Then
            Keyword = Web.HttpUtility.UrlDecode(Request.QueryString("keyword"))
        Else
            Keyword = ""
        End If
        '-----------------------------------------------------------------------------
        TabText.Text = WebUtility.GetMyAreaLinks(MyAreaLink.myknowledge_trace_lp)
        '-----------------------------------------------------------------------------

        Dim link1 As String = "/knowledge/myknowledge_question_lp.aspx?ArticleType={0}&CategoryId=" & CategoryId & "&keyword=" & Web.HttpUtility.UrlEncode(Keyword)
        Dim link2 As String = ""
        Dim imagestr As String = ""

        If Sort = "D" Then
            link2 = link1 & "&Sort=A"
            imagestr = "<img src=""images/down.gif"" alt=""由新到舊"" border=""0"" />"
        ElseIf Sort = "A" Then
            link2 = link1 & "&Sort=D"
            imagestr = "<img src=""images/up.gif"" alt=""由舊到新"" border=""0"" />"
        End If

        sb = New StringBuilder()

        If ArticleType = "A" Then
            sb.Append("建立日期<a href=""" & link2.Replace("{0}", "A") & """>" & imagestr & "</a>│<a href=""" & link1.Replace("{0}", "B") & """>更新日期</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "C") & """>討論數</a>│<a href=""" & link1.Replace("{0}", "D") & """>評價</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "E") & """>瀏覽數</a>│<a href=""" & link1.Replace("{0}", "F") & """>專家補充</a>")
        ElseIf ArticleType = "B" Then
            sb.Append("<a href=""" & link1.Replace("{0}", "A") & """>建立日期</a>│更新日期<a href=""" & link2.Replace("{0}", "B") & """>" & imagestr & "</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "C") & """>討論數</a>│<a href=""" & link1.Replace("{0}", "D") & """>評價</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "E") & """>瀏覽數</a>│<a href=""" & link1.Replace("{0}", "F") & """>專家補充</a>")
        ElseIf ArticleType = "C" Then
            sb.Append("<a href=""" & link1.Replace("{0}", "A") & """>建立日期</a>│<a href=""" & link1.Replace("{0}", "B") & """>更新日期</a>" & _
                      "│討論數<a href=""" & link2.Replace("{0}", "C") & """>" & imagestr & "</a>│<a href=""" & link1.Replace("{0}", "D") & """>評價</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "E") & """>瀏覽數</a>│<a href=""" & link1.Replace("{0}", "F") & """>專家補充</a>")
        ElseIf ArticleType = "D" Then
            sb.Append("<a href=""" & link1.Replace("{0}", "A") & """>建立日期</a>│<a href=""" & link1.Replace("{0}", "B") & """>更新日期</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "C") & """>討論數</a>│評價<a href=""" & link2.Replace("{0}", "D") & """>" & imagestr & "</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "E") & """>瀏覽數</a>│<a href=""" & link1.Replace("{0}", "F") & """>專家補充</a>")
        ElseIf ArticleType = "E" Then
            sb.Append("<a href=""" & link1.Replace("{0}", "A") & """>建立日期</a>│<a href=""" & link1.Replace("{0}", "B") & """>更新日期</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "C") & """>討論數</a>│<a href=""" & link1.Replace("{0}", "D") & """>評價</a>" & _
                      "│瀏覽數<a href=""" & link2.Replace("{0}", "E") & """>" & imagestr & "</a>│<a href=""" & link1.Replace("{0}", "F") & """>專家補充</a>")
        ElseIf ArticleType = "F" Then
            sb.Append("<a href=""" & link1.Replace("{0}", "A") & """>建立日期</a>│<a href=""" & link1.Replace("{0}", "B") & """>更新日期</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "C") & """>討論數</a>│<a href=""" & link1.Replace("{0}", "D") & """>評價</a>" & _
                      "│<a href=""" & link1.Replace("{0}", "E") & """>瀏覽數</a>│專家補充<a href=""" & link2.Replace("{0}", "F") & """>" & imagestr & "</a>")
        End If

        ArticleTypeText.Text = sb.ToString()
        sb = Nothing
        '-----------------------------------------------------------------------------

        Dim PageNumber As Integer
        Dim PageSize As Integer

        If Not Page.IsPostBack Then
            Try
                If Request.QueryString("PageSize") = "" Then
                    PageSize = 10
                Else
                    PageSize = Integer.Parse(Request.QueryString("PageSize"))
                End If
                If Request.QueryString("PageNumber") = "" Then
                    PageNumber = 1
                Else
                    PageNumber = Integer.Parse(Request.QueryString("PageNumber"))
                End If
            Catch ex As Exception
                If Request.QueryString("debug") = "true" Then
                    Response.Write(ex.ToString)
                    Response.End()
                End If
                Response.Redirect("/knowledge/myknowledge_trace_lp.aspx")
                Response.End()
            End Try
        Else
            PageSize = PageSizeDDL.SelectedValue
            If PageNumberDDL.SelectedValue = "" Then
                PageNumber = 1
            Else
                PageNumber = PageNumberDDL.SelectedValue
            End If
            CategoryId = QuestionTypeDDL.SelectedValue
            Keyword = ""
        End If
        '-----------------------------------------------------------------------------

        If CategoryId = "" Then
            QuestionTypeDDL.SelectedIndex = 0
        ElseIf CategoryId = "A" Then
            QuestionTypeDDL.SelectedIndex = 1
        ElseIf CategoryId = "B" Then
            QuestionTypeDDL.SelectedIndex = 2
        ElseIf CategoryId = "C" Then
            QuestionTypeDDL.SelectedIndex = 3
        ElseIf CategoryId = "D" Then
            QuestionTypeDDL.SelectedIndex = 4
        ElseIf CategoryId = "E" Then
            QuestionTypeDDL.SelectedIndex = 5
        End If
        '-----------------------------------------------------------------------------

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim myTable As DataTable
        Dim myDataRow As DataRow

        Dim Total As Integer = 0
        Dim PageCount As Integer = 0
        Dim Position As Integer = 1

        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.fCTUPublic, CuDTGeneric.sTitle, CodeMain.mValue, CuDTGeneric.xPostDate, CuDTGeneric.dEditDate, CuDTGeneric.topCat, "
            sqlString &= "KnowledgeForum.DiscussCount, KnowledgeForum.BrowseCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, CuDTGeneric_1.iCUItem AS iCUItem_1, "
            sqlString &= "KnowledgeForum.HavePros FROM CuDTGeneric AS CuDTGeneric_1 INNER JOIN CuDTGeneric INNER JOIN KnowledgeForum "
            sqlString &= "ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN KnowledgeForum AS KnowledgeForum_1 ON "
            sqlString &= "CuDTGeneric.iCUItem = KnowledgeForum_1.ParentIcuitem ON CuDTGeneric_1.iCUItem = KnowledgeForum_1.gicuitem INNER JOIN "
            sqlString &= "CodeMain ON CuDTGeneric.topCat = CodeMain.mCode WHERE (CuDTGeneric_1.iCTUnit = @ictunit) AND (CodeMain.codeMetaID = 'KnowledgeType') "
            sqlString &= "AND (CuDTGeneric_1.iEditor = @account) AND (CuDTGeneric_1.siteId = @siteid) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric_1.fCTUPublic = 'Y') "
            sqlString &= "AND (KnowledgeForum.Status = 'N') AND (KnowledgeForum_1.Status = 'N') "

            '---農林漁牧其中一種---
            If CategoryId <> "" Then
                sqlString &= "AND CuDTGeneric.topCat = '" & CategoryId & "' "
            End If
            If Keyword <> "" Then
                sqlString &= "AND CuDTGeneric.sTitle LIKE '%" & Keyword & "%' "
            End If
            '---排序種類---
            If ArticleType = "A" Then
                sqlString &= "ORDER BY CuDTGeneric.xPostDate "
            ElseIf ArticleType = "B" Then
                sqlString &= "ORDER BY CuDTGeneric.dEditDate "
            ElseIf ArticleType = "C" Then
                sqlString &= "ORDER BY KnowledgeForum.DiscussCount "
            ElseIf ArticleType = "D" Then
                sqlString &= "ORDER BY KnowledgeForum.GradeCount "
            ElseIf ArticleType = "E" Then
                sqlString &= "ORDER BY KnowledgeForum.BrowseCount "
            ElseIf ArticleType = "F" Then
                sqlString &= "AND KnowledgeForum.HavePros = 'Y' ORDER BY CuDTGeneric.xPostDate "
            End If
            If Sort = "D" Then
                sqlString &= "DESC"
            End If

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)

            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeTraceCtUnitId"))
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
            myCommand.Parameters.AddWithValue("@account", MemberId)

            myReader = myCommand.ExecuteReader()

            myTable = New DataTable()
            myTable.Load(myReader)

            Total = myTable.Rows().Count()

            DeleteTraceBtn.OnClientClick = "CheckDeleteBtn(" & Total & ")"

            If Total <> 0 Then

                PageCount = Int(Total / PageSize + 0.999)

                If PageCount < PageNumber Then
                    PageNumber = PageCount
                End If

                PageNumberText.Text = PageNumber.ToString()
                TotalPageText.Text = PageCount.ToString()
                TotalRecordText.Text = Total.ToString()
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
                For j = 0 To PageCount - 1
                    item = New ListItem
                    item.Value = j + 1
                    item.Text = j + 1
                    If PageNumber = (j + 1) Then
                        item.Selected = True
                    End If
                    PageNumberDDL.Items.Insert(j, item)
                    item = Nothing
                Next

                Position = PageSize * (PageNumber - 1)

                myDataRow = myTable.Rows(Position)

                If PageNumber > 1 Then
                    PreviousText.Visible = True
                    PreviousImg.Visible = True
                    PreviousLink.NavigateUrl = "/knowledge/myknowledge_trace_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & _
                                               "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                Else
                    PreviousText.Visible = False
                    PreviousImg.Visible = False
                End If
                If PageNumberDDL.SelectedValue < PageCount Then
                    NextText.Visible = True
                    NextImg.Visible = True
                    NextLink.NavigateUrl = "/knowledge/myknowledge_trace_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & _
                                           "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                Else
                    NextText.Visible = False
                    NextImg.Visible = False
                End If

                Dim i As Integer = 0
                Dim link As String = ""

                sb = New StringBuilder
                sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"">")
                sb.Append("<tr><th scope=""col"">刪除</th><th scope=""col"">狀態</th><th scope=""col"">問題標題</th><th scope=""col"">知識分類</th><th scope=""col"">建立日期</th>")
                sb.Append("<th scope=""col"">更新日期</th><th scope=""col"">討論</th><th scope=""col"">評價</th><th scope=""col"">瀏覽</th><th scope=""col"">專家</th></tr>")

                Dim totalDiscuss As String
                For i = 0 To PageSize - 1

                    sb.Append("<tr>")

                    '---checkbox---
                    sb.Append("<td><input type=""checkbox"" name=""tracecheckbox" & i & """ value=""" & myDataRow("iCUItem_1") & """ /></td>")

                    '---state---
                    sb.Append("<td>&nbsp;</td>")
                    link = ""

                    link = "<a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myDataRow("iCUItem") & "&ArticleType=A&CategoryId=" & myDataRow("topCat") & """>"
                    link &= myDataRow("sTitle").ToString() + "</a>"

                    '計算討論數
                    totalDiscuss = "0"
                    Dim knowledgeUtility As New KnowledgeUtility(myDataRow("iCUItem"))
                    totalDiscuss = knowledgeUtility.GetDiscussCount()

                    sb.Append("<td>" & link & "</td>")
                    sb.Append("<td>" & myDataRow("mValue") & "</td>")
                    sb.Append("<td>[" & Date.Parse(myDataRow("xPostDate")).ToShortDateString & "]</td>")
                    sb.Append("<td>[" & Date.Parse(myDataRow("dEditDate")).ToShortDateString & "]</td>")
                    'sb.Append("<td>" & myDataRow("DiscussCount") & "</td>")
                    sb.Append("<td>" & totalDiscuss & "</td>")

                    sb.Append("<td><div class=""staricon"">")
                    sb.Append(ConcatStarString(Integer.Parse(myDataRow("GradeCount")), Integer.Parse(myDataRow("GradePersonCount"))))
                    sb.Append("</div></td>")

                    sb.Append("<td>" & myDataRow("BrowseCount") & "</td>")
                    If myDataRow("HavePros") = "Y" Then
                        sb.Append("<td>V</td>")
                    Else
                        sb.Append("<td>&nbsp;</td>")
                    End If
                    sb.Append("</tr>")

                    Position += 1
                    If myTable.Rows.Count <= Position Then
                        Exit For
                    End If
                    myDataRow = myTable.Rows(Position)
                Next

                sb.Append("</table>")

                TableText.Text = sb.ToString()

            End If

        Catch ex As Exception
            If Request("debug") = "true" Then
                Response.Write(ex.ToString())
                Response.End()
            End If
            Response.Redirect("/knowledge/knowledge.aspx")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Sub

    Protected Sub SearchBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SearchBtn.Click

        Dim keyword As String = Web.HttpUtility.UrlEncode(KeywordText.Text.Trim)

        Response.Redirect("/knowledge/myknowledge_trace_lp.aspx?keyword=" & keyword)

    End Sub

    Protected Sub DeleteTraceBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteTraceBtn.Click

        If DeleteItem.Text.Length > 0 Then

            Dim MemberId As String
            Dim Script1 As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>"
            '-----------------------------------------------------------------------------
            If Session("memID") = "" Or Session("memID") Is Nothing Then
                MemberId = ""
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script1)
                Exit Sub
            Else
                MemberId = Session("memID").ToString()
            End If
            '-----------------------------------------------------------------------------

            Dim Script As String = "<script>alert('{0}!!');window.location.href='" & Request.RawUrl() & "';</script>"

            Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
            Dim sqlString As String = ""
            Dim myConnection As SqlConnection
            Dim myCommand As SqlCommand
            Dim icuitem As String = DeleteItem.Text.Trim.Substring(0, DeleteItem.Text.Trim.Length - 1)

            myConnection = New SqlConnection(ConnString)
            Try
                sqlString = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem IN (" & icuitem & ")"
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.ExecuteNonQuery()
                myCommand.Dispose()

                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "DeleteScript", Script.Replace("{0}", "成功刪除我的追蹤"))

            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/knowledge/myknowledge_trace_lp.aspx")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try
        End If

    End Sub

    Function ConcatStarString(ByVal GradeCount As Integer, ByVal PersonCount As Integer) As String

        Dim dresult As Double = 0

        If PersonCount = 0 Then
            dresult = GradeCount / 1
        Else
            dresult = GradeCount / PersonCount
        End If

        Dim str As String = ""
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

        Return str

    End Function

End Class