Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Member_MemberRanking
    Inherits System.Web.UI.Page

  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sqlString As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader
  Dim myTable As DataTable
    Dim myDataRow As DataRow
    Dim MemberId As String

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim cat As String = "A"

        If Request.QueryString("cat") <> "A" And Request.QueryString("cat") <> "B" And Request.QueryString("cat") <> "C" And Request.QueryString("cat") <> "D" And Request.QueryString("cat") <> "MyRank" Then
            Response.Redirect("/Member/MemberRankingThisYear.aspx?cat=A")
            Response.End()
        Else
            cat = Request.QueryString("cat")
        End If

        '---list---
        Dim PageNumber As Integer
        Dim PageSize As Integer

        If Not Page.IsPostBack Then
            Try
                If Request.QueryString("PageSize") = "" Then
                    PageSize = 15
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
                Response.Redirect("/")
                Response.End()
            End Try
        Else
            PageSize = PageSizeDDL.SelectedValue
            PageNumber = PageNumberDDL.SelectedValue
        End If

        Dim gradeLevel1 As Integer = 0
        Dim gradeLevel2 As Integer = 0
        Dim gradeLevel3 As Integer = 0
        Dim gradeLevel4 As Integer = 0

        '---取出分數上限---
        GetGradeLevel(gradeLevel1, gradeLevel2, gradeLevel3, gradeLevel4)

        GetTabTextByType(cat, gradeLevel1, gradeLevel2, gradeLevel3, gradeLevel4)

        Dim Total As Integer = 0
        Dim PageCount As Integer = 0
        Dim Position As Integer = 1

        myConnection = New SqlConnection(ConnString)
        Try
            myConnection.Open()
            If cat = "MyRank" And String.IsNullOrEmpty(Session("memID")) <> True Then
                GetPersonalRank(MemberId)
                TabText.Text = "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=A"">積分總排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=B"">成就排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=C"">貢獻排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=D"">精神排行</a></li>"
                TabText.Text &= "<li class=""active""><a href=""/Member/MemberRankingThisYear.aspx?cat=MyRank"">我的排行</a></li>" '2011.07.11 Grace

            ElseIf String.IsNullOrEmpty(Session("memID")) = True And cat = "MyRank" Then 'Grace 這邊導頁可能有問題
                Response.Redirect("/Member/MemberRankingThisYear.aspx?cat=A")
                Response.End()

            ElseIf cat <> "MyRank" Then

                'Bob 2009/08/28
                sqlString = ""
                sqlString &= "SELECT "
                sqlString &= "Member.account"
                sqlString &= ", Member.realname"
                sqlString &= ", Member.nickname"
                sqlString &= ", isnull(GradeBrowse_thisyear.GradeBrowse, 0) as browseTotal"
                sqlString &= ", isnull(GradeLogin_thisyear.GradeLogin, 0) as loginTotal"
                sqlString &= ", isnull(GradeShare_thisyear.GradeShare, 0) as shareTotal , isnull(B.calculateTotal,0) as allyearstotal "
                sqlString &= ", (isnull(GradeBrowse_thisyear.GradeBrowse, 0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin, 0) * 0.2) + isnull(GradeShare_thisyear.GradeShare, 0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) * 0.2 + isnull(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 AS calculateTotal "
                sqlString &= "FROM Member left JOIN GradeLogin_thisyear ON Member.account = GradeLogin_thisyear.memberId "
                sqlString &= "left JOIN GradeShare_thisyear ON Member.account = GradeShare_thisyear.memberId "
                sqlString &= "left JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId "
                sqlString &= "left JOIN MemberGradeSummary B ON Member.account = B.memberId  "
                sqlString &= " LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor  "
                sqlString &= " LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor  "
                sqlString &= " LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor "
                sqlString &= " LEFT JOIN View_MemberGradeAdditional_ThisYear VIEWADDTHIS ON Member.account = VIEWADDTHIS.memberId "
                sqlString &= " LEFT JOIN dbo.MemberGradeContentByYear MGCBY ON Member.account = MGCBY.memberId AND (CONVERT(INT, CONVERT(varchar(4), getdate(), 120 ) ) - ISNULL(years,0) = 0) "
                Select Case cat
                    Case "A"
                        sqlString &= "ORDER BY calculateTotal DESC "
                    Case "B"
                        sqlString &= "ORDER BY browseTotal DESC "
                    Case "C"
                        sqlString &= "ORDER BY shareTotal DESC "
                    Case "D"
                        sqlString &= "ORDER BY loginTotal DESC "
                End Select


                myCommand = New SqlCommand(sqlString, myConnection)
                myReader = myCommand.ExecuteReader()

                myTable = New DataTable()
                myTable.Load(myReader)

                Total = myTable.Rows().Count()
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

                If Total <> 0 Then

                    PageCount = Int(Total / PageSize + 0.999)

                    If PageCount < PageNumber Then
                        PageNumber = PageCount
                    End If

                    PageNumberText.Text = PageNumber.ToString()
                    TotalPageText.Text = PageCount.ToString()
                    TotalRecordText.Text = Total.ToString()
                    If PageSize = 15 Then
                        PageSizeDDL.SelectedIndex = 0
                    ElseIf PageSize = 30 Then
                        PageSizeDDL.SelectedIndex = 1
                    ElseIf PageSize = 50 Then
                        PageSizeDDL.SelectedIndex = 2
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

                    myDataRow = myTable.Rows.Item(Position)

                    If PageNumber > 1 Then
                        PreviousLink.NavigateUrl = "/Member/MemberRankingThisYear.aspx?cat=" & cat & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                    Else
                        PreviousText.Enabled = False
                    End If
                    If PageNumberDDL.SelectedValue < PageCount Then
                        NextLink.NavigateUrl = "/Member/MemberRankingThisYear.aspx?cat=" & cat & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                    Else
                        NextText.Enabled = False
                    End If

                    Dim i As Integer = 0
                    Dim link As String = ""
                    Dim sb As New StringBuilder

                    sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"">")

                    If cat = "A" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">總得點 (得點*百分比(權重)之加總)</th>")
                        sb.Append("</tr>")
                    ElseIf cat = "B" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">網站瀏覽得點 (瀏覽原始得點)</th>")
                        sb.Append("</tr>")
                    ElseIf cat = "C" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">參與互動得點 (互動原始得點)</th>")
                        sb.Append("</tr>")
                    ElseIf cat = "D" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">使用頻率得點 (登入原始得點)</th>")
                        sb.Append("</tr>")
                    End If

                    For i = 0 To PageSize - 1

                        sb.Append("<tr><td>" & (PageSize * (PageNumber - 1)) + i + 1 & "</td>")

                        If myDataRow("allyearstotal") >= gradeLevel4 Then
                            sb.Append("<td>達人級會員</td>")
                        ElseIf myDataRow("allyearstotal") >= gradeLevel3 Then
                            sb.Append("<td>高手級會員</td>")
                        ElseIf myDataRow("allyearstotal") >= gradeLevel2 Then
                            sb.Append("<td>進階級會員</td>")
                        ElseIf myDataRow("allyearstotal") >= gradeLevel1 Then
                            sb.Append("<td>入門級會員</td>")
                        End If

                        sb.Append("<td>")

                        Dim realName As String = Server.HtmlDecode(myDataRow("realname").ToString)
                        If myDataRow("nickname") IsNot DBNull.Value Then
                            If myDataRow("nickname") <> "" Then
                                sb.Append(myDataRow("nickname"))
                            Else
                                If myDataRow("realname") IsNot DBNull.Value Then
                                    If realName <> "" Then
                                        sb.Append((realName.Substring(0, 1) & "＊" & realName.Substring(myDataRow("realname").ToString.Trim.Length - 1, 1)))
                                    End If
                                End If
                            End If
                        Else
                            If myDataRow("realname") IsNot DBNull.Value Then
                                If realName <> "" Then
                                    sb.Append((realName.Substring(0, 1) & "＊" & realName.Substring(realName.Trim.Length - 1, 1)))
                                End If
                            End If
                        End If
                        sb.Append("</td>")

                        If cat = "A" Then
                            sb.Append("<td>" & myDataRow("calculateTotal") & "</td>")
                        ElseIf cat = "B" Then
                            sb.Append("<td>" & myDataRow("browseTotal") & "</td>")
                        ElseIf cat = "C" Then
                            sb.Append("<td>" & myDataRow("shareTotal") & "</td>")
                        ElseIf cat = "D" Then
                            sb.Append("<td>" & myDataRow("loginTotal") & "</td>")
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

                Else

                    Dim sb As New StringBuilder

                    PageNumberText.Text = "0"
                    TotalPageText.Text = "0"
                    TotalRecordText.Text = "0"

                    sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"">")

                    If cat = "A" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">總得點</th>")
                        sb.Append("</tr>")
                    ElseIf cat = "B" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">網站瀏覽得點</th>")
                        sb.Append("</tr>")
                    ElseIf cat = "C" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">參與互動得點</th>")
                        sb.Append("</tr>")
                    ElseIf cat = "D" Then
                        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                        sb.Append("<th scope=""col"">等級</th>")
                        sb.Append("<th scope=""col"">會員</th>")
                        sb.Append("<th scope=""col"">使用頻率得點</th>")
                        sb.Append("</tr>")
                    End If
                    sb.Append("</table>")

                    TableText.Text = sb.ToString

                End If
            End If


        Catch ex As Exception
            If Request("debug") = "true" Then
                Response.Write(ex.ToString())
                Response.End()
            End If
            Response.Redirect("/")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Sub

    Sub GetGradeLevel(ByRef level1 As Integer, ByRef level2 As Integer, ByRef level3 As Integer, ByRef level4 As Integer)

        myConnection = New SqlConnection(ConnString)
        Try
            myConnection.Open()
            sqlString = "SELECT * FROM CodeMain WHERE CodeMetaID = 'gradeLevel'"
            myCommand = New SqlCommand(sqlString, myConnection)
            myReader = myCommand.ExecuteReader
            While myReader.Read
                If myReader("mCode") = "1" Then level1 = myReader("mValue")
                If myReader("mCode") = "2" Then level2 = myReader("mValue")
                If myReader("mCode") = "3" Then level3 = myReader("mValue")
                If myReader("mCode") = "4" Then level4 = myReader("mValue")
            End While
            If Not myReader.IsClosed Then myReader.Close()
            myCommand.Dispose()
        Catch ex As Exception
        Finally
            If myConnection.State = ConnectionState.Open Then myConnection.Close()
        End Try

    End Sub

    Sub GetTabTextByType(ByVal cat As String, ByVal gradeLevel1 As Integer, ByVal gradeLevel2 As Integer, ByVal gradeLevel3 As Integer, ByVal gradeLevel4 As Integer)

        myConnection = New SqlConnection(ConnString)
        Try
            myConnection.Open()

            If cat = "A" Then

                Dim level11 As Integer = 0, level12 As Integer = 0, level13 As Integer = 0, level14 As Integer = 0
                Dim level21 As Integer = 0, level22 As Integer = 0

                sqlString = "SELECT COUNT(*) FROM MemberGradeSummary WHERE calculateTotal > @gradelevel1 AND calculateTotal <  @gradeLevel2"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@gradelevel1", gradeLevel1)
                myCommand.Parameters.AddWithValue("@gradelevel2", gradeLevel2)
                level11 = myCommand.ExecuteScalar
                myCommand.Dispose()

                sqlString = "SELECT COUNT(*) FROM MemberGradeSummary WHERE calculateTotal >= @gradelevel2 AND calculateTotal <  @gradeLevel3"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@gradelevel2", gradeLevel2)
                myCommand.Parameters.AddWithValue("@gradeLevel3", gradeLevel3)
                level12 = myCommand.ExecuteScalar
                myCommand.Dispose()

                sqlString = "SELECT COUNT(*) FROM MemberGradeSummary WHERE calculateTotal >= @gradelevel3 AND calculateTotal <  @gradeLevel4"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@gradelevel3", gradeLevel3)
                myCommand.Parameters.AddWithValue("@gradeLevel4", gradeLevel4)
                level13 = myCommand.ExecuteScalar
                myCommand.Dispose()

                sqlString = "SELECT COUNT(*) FROM MemberGradeSummary WHERE calculateTotal >= @gradelevel4"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@gradelevel4", gradeLevel4)
                level14 = myCommand.ExecuteScalar
                myCommand.Dispose()

                level21 = level13 + level14
                level22 = level12 + level13 + level14

                browseby.Text = "<p>會員總數：【入門級會員 <span class=""number2"">" & level11 & "</span>人】 【進階級會員 <span class=""number2"">" & level12 & "</span>人】"
                browseby.Text &= " 【高手級會員 <span class=""number2"">" & level13 & "</span>人】 【達人級會員 <span class=""number2"">" & level14 & "</span>人】</p>"
                'browseby.Text &= "<p>年度積分活動全能獎 【42吋數位液晶顯示器】：1 獎項 / <span class=""number2"">" & level21 & "</span>人 (等級達「高手級」以上之所有會員)<br/>"
                'browseby.Text &= "年度積分活知識獎 【Wii 中文版遊戲機】：2 獎項 / <span class=""number2"">" & level22 & "</span>人 (等級達「進階級」以上之所有會員)</p>"
                TabText.Text = "<li class=""active""><a href=""/Member/MemberRankingThisYear.aspx?cat=A"">積分總排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=B"">成就排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=C"">貢獻排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=D"">精神排行</a></li>"
                If String.IsNullOrEmpty(Session("memID")) <> True Then
                    TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=MyRank"">我的排行</a></li>" '2011.07.11 Grace
                End If

            ElseIf cat = "B" Then

                Dim level1 As Integer = 0
                sqlString = "SELECT COUNT(*) FROM MemberGradeSummary WHERE browseTotal >= 150"
                myCommand = New SqlCommand(sqlString, myConnection)
                level1 = myCommand.ExecuteScalar
                myCommand.Dispose()
                'browseby.Text = "<p>年度積分活動成就獎 【2.5吋 250G 行動硬碟】：3 獎項 / <span class=""number2"">" & level1 & "</span>人 (瀏覽(原始)得點，高於150(含)點以上者)</p>"
                TabText.Text = "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=A"">積分總排行</a></li>"
                TabText.Text &= "<li class=""active""><a href=""/Member/MemberRankingThisYear.aspx?cat=B"">成就排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=C"">貢獻排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=D"">精神排行</a></li>"
                If String.IsNullOrEmpty(Session("memID")) <> True Then
                    TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=MyRank"">我的排行</a></li>" '2011.07.11 Grace
                End If

            ElseIf cat = "C" Then

                Dim level1 As Integer = 0
                sqlString = "SELECT COUNT(*) FROM MemberGradeSummary WHERE shareTotal >= 150"
                myCommand = New SqlCommand(sqlString, myConnection)
                level1 = myCommand.ExecuteScalar
                myCommand.Dispose()
                ' browseby.Text = "<p>年度積分活動貢獻獎 【8G隨身碟】：20 獎項 / <span class=""number2"">" & level1 & "</span>人 (互動(原始)得點，高於150(含)點以上者)</p>"
                TabText.Text = "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=A"">積分總排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=B"">成就排行</a></li>"
                TabText.Text &= "<li class=""active""><a href=""/Member/MemberRankingThisYear.aspx?cat=C"">貢獻排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=D"">精神排行</a></li>"
                If String.IsNullOrEmpty(Session("memID")) <> True Then
                    TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=MyRank"">我的排行</a></li>" '2011.07.11 Grace
                End If

            ElseIf cat = "D" Then

                Dim level1 As Integer = 0
                sqlString = "SELECT COUNT(*) FROM MemberGradeSummary WHERE loginTotal >= 150"
                myCommand = New SqlCommand(sqlString, myConnection)
                level1 = myCommand.ExecuteScalar
                myCommand.Dispose()
                'browseby.Text = "<p>年度積分活動精神獎 【圖書禮券 1,000元】：10 獎項 / <span class=""number2"">" & level1 & "</span>人 (登入(原始)得點，高於150(含)點以上者)</p>"
                TabText.Text = "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=A"">積分總排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=B"">成就排行</a></li>"
                TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=C"">貢獻排行</a></li>"
                TabText.Text &= "<li class=""active""><a href=""/Member/MemberRankingThisYear.aspx?cat=D"">精神排行</a></li>"
                If String.IsNullOrEmpty(Session("memID")) <> True Then
                    TabText.Text &= "<li><a href=""/Member/MemberRankingThisYear.aspx?cat=MyRank"">我的排行</a></li>" '2011.07.11 Grace
                End If
           
            End If

        Catch ex As Exception
        Finally
            If myConnection.State = ConnectionState.Open Then myConnection.Close()
        End Try

    End Sub


    Sub GetPersonalRank(ByVal memberId As String)
        'Function GetPersonalRank(ByVal memberId As String) As String


        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlStr As String = ""
        Dim str As String = ""
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim sb As New StringBuilder

        Dim grade1 As Integer = 0, grade2 As Integer = 0, grade3 As Integer = 0, grade4 As Integer = 0
        Dim rank1 As Integer = 0, rank2 As Integer = 0, rank3 As Integer = 0, rank4 As Integer = 0
        sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格""><tr>")
        If String.IsNullOrEmpty(Session("memID")) <> True Then

            sqlStr = ""
            sqlStr = "WITH subjectOrder AS ( select row_number() over ( ORDER BY (isnull(GradeBrowse_thisyear.GradeBrowse, 0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin, 0) * 0.2) + isnull(GradeShare_thisyear.GradeShare, 0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) * 0.2 + isnull(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 DESC ) as rowid , B.memberId , isnull(GradeBrowse_thisyear.GradeBrowse, 0) as browseTotal , isnull(GradeLogin_thisyear.GradeLogin, 0) as loginTotal , isnull(GradeShare_thisyear.GradeShare, 0) as shareTotal , isnull(B.calculateTotal,0) as allyearstotal , (isnull(GradeBrowse_thisyear.GradeBrowse, 0) * 0.15 + isnull(GradeLogin_thisyear.GradeLogin, 0) * 0.2) + isnull(GradeShare_thisyear.GradeShare, 0) * 0.3 + (ISNULL(VCB.ThisYearBrowseGrade,0)+ISNULL(VCC.ThisYearCommendGrade,0)+ISNULL(VCD.ThisYearDiscussGrade,0)+ISNULL(MGCBY.contentIQPoint,0)+ISNULL(MGCBY.contentRaisePoint,0)+ISNULL(MGCBY.contentLogoPoint,0)+ ISNULL(MGCBY.contentPlantLogPoint,0)+ISNULL(MGCBY.contentFishBowlPoint,0)+ ISNULL(MGCBY.contentDrPoint,0)+ ISNULL(MGCBY.contentKnowledgeQAPoint,0)+ ISNULL(MGCBY.contentDr2010Point,0)+ISNULL(MGCBY.contentTreasureHuntPoint,0)) * 0.2 + isnull(VIEWADDTHIS.ThisYearADDpoint, 0) * 0.15 AS calculateTotal FROM Member left JOIN GradeLogin_thisyear ON Member.account = GradeLogin_thisyear.memberId left JOIN GradeShare_thisyear ON Member.account = GradeShare_thisyear.memberId left JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId left JOIN MemberGradeSummary B ON Member.account = B.memberId LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor LEFT JOIN View_MemberGradeAdditional_ThisYear VIEWADDTHIS ON Member.account = VIEWADDTHIS.memberId LEFT JOIN dbo.MemberGradeContentByYear MGCBY ON Member.account = MGCBY.memberId AND (CONVERT(INT, CONVERT(varchar(4), getdate(), 120 ) ) - ISNULL(years,0) = 0) ) SELECT * FROM subjectOrder WHERE memberId = @memberid"
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", Session("memID"))
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                grade1 = myReader("calculateTotal")
                rank1 = myReader("rowid")
            End If
            myReader.Close()
            myCommand.Dispose()

            sqlStr = ""
            sqlStr = "WITH subjectOrder AS (select row_number() over(ORDER BY isnull(GradeBrowse_thisyear.GradeBrowse, 0) DESC)as rowid, GradeBrowse_thisyear.memberId, isnull(GradeBrowse_thisyear.GradeBrowse, 0) as browseTotal  FROM Member left JOIN GradeBrowse_thisyear ON Member.account = GradeBrowse_thisyear.memberId LEFT JOIN GradeContentBrowse_thisyear VCB ON Member.account = VCB.ieditor)SELECT * FROM subjectOrder WHERE memberId = @memberid "
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", Session("memID"))
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                grade2 = myReader("browseTotal")
                rank2 = myReader("rowid")
            End If
            myReader.Close()
            myCommand.Dispose()

            sqlStr = ""
            sqlStr = "WITH subjectOrder AS ( select row_number() over(ORDER BY isnull(GradeShare_thisyear.GradeShare, 0) DESC)as rowid , GradeShare_thisyear.memberId , isnull(GradeShare_thisyear.GradeShare, 0) as shareTotal FROM Member left JOIN GradeShare_thisyear ON Member.account = GradeShare_thisyear.memberId LEFT JOIN GradeContentCommend_thisyear VCC ON Member.account = VCC.ieditor ) SELECT * FROM subjectOrder WHERE memberId = @memberid		"
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", Session("memID"))
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                grade3 = myReader("shareTotal")
                rank3 = myReader("rowid")
            End If
            myReader.Close()
            myCommand.Dispose()

            sqlStr = ""
            sqlStr = "WITH subjectOrder AS ( select row_number() over(ORDER BY isnull(GradeLogin_thisyear.GradeLogin, 0) DESC)as rowid , GradeLogin_thisyear.memberId , isnull(GradeLogin_thisyear.GradeLogin, 0) as loginTotal FROM Member left JOIN GradeLogin_thisyear ON Member.account = GradeLogin_thisyear.memberId LEFT JOIN GradeContentDiscuss_thisyear VCD ON Member.account = VCD.ieditor ) SELECT * FROM subjectOrder WHERE memberId = @memberid"
            myCommand = New SqlCommand(sqlStr, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", Session("memID"))
            myReader = myCommand.ExecuteReader
            If myReader.Read Then
                grade4 = myReader("loginTotal")
                rank4 = myReader("rowid")
            End If
            myReader.Close()
            myCommand.Dispose()

            sb.Append("<center><table width=""100%"" cellspacing=""0"" cellpadding=""0"" >")
            sb.Append("<tr><th scope=""col"" ><B>排行類別</B> </th><th  scope=""col"" ><B>積分</B></th><th scope=""col"" ><B>排名</B></th></tr>")
            sb.Append("<tr><td scope=""col"" ALIGN=""center"" width='35%'><B>積分總排行</B> </td><td scope=""col"" ALIGN=""center""><span>" & grade1 & "</span></td><td scope=""col"" ALIGN=""center"">" & rank1 & "</td></tr>")
            sb.Append("<tr><td scope=""col"" ALIGN=""center""><B>成就排行</B></td><td scope=""col"" ALIGN=""center""><span>" & grade2 & "</span></td><td scope=""col"" ALIGN=""center"">" & rank2 & "</td></tr>")
            sb.Append("<tr><td scope=""col"" ALIGN=""center""><B>貢獻排行</B></td><td  scope=""col"" ALIGN=""center""><span>" & grade3 & "</span></td><td scope=""col"" ALIGN=""center"">" & rank3 & "</td></tr>")
            sb.Append("<tr><td scope=""col"" ALIGN=""center""><B>精神排行</B></td><td scope=""col"" ALIGN=""center""><span>" & grade4 & "</span></td><td scope=""col"" ALIGN=""center"">" & rank4 & "</td></tr>")
            sb.Append("</tabe></center>")
        End If
        sb.Append("</tr></table>")
        TableText.Text = sb.ToString

    End Sub

End Class
