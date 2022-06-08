Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_discuss_lp
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ArticleId As String = ""
    Dim ArticleType As String = ""
    Dim CategoryId As String = ""
    Dim PageNumber As Integer = 1
    Dim PageSize As Integer = 10
    Dim MemberId As String = ""
    Dim MemberActor As String = ""

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
      Try
        ArticleId = Integer.Parse(Request.QueryString("ArticleId"))
      Catch ex As Exception
        Response.Redirect("/knowledge/knowledge.aspx")
        Response.End()
      End Try
    End If

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
        Dim browse As New KPIBrowse(Session("memID"), "browseDiscussLP", ArticleId)
        browse.HandleBrowse()
      End If
      '---end of kpi use---
      Dim relink As String = "/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&kpi=0"
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
    BackText.Text = "<a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """ class=""Back"">回上一頁</a>"

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
        Response.Redirect("/knowledge/knowledge.aspx")
        Response.End()
      End Try
    Else
      PageSize = PageSizeDDL.SelectedValue
      PageNumber = PageNumberDDL.SelectedValue
    End If

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim sb As StringBuilder
    Dim myTable As DataTable
    Dim myDataRow As DataRow
    Dim Total As Integer = 0
    Dim PageCount As Integer = 0
    Dim Position As Integer = 1
    Dim DiscussFlag As Boolean = True

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xBody, CuDTGeneric.xPostDate, CuDTGeneric.xNewWindow FROM KnowledgeForum INNER JOIN CuDTGeneric ON "
      sqlString &= "KnowledgeForum.gicuitem = CuDTGeneric.iCUItem WHERE (CuDTGeneric.iCUItem = @ArticleId) AND (CuDTGeneric.fCTUPublic = 'Y') AND (KnowledgeForum.Status = 'N')"
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ArticleId", ArticleId)
      myReader = myCommand.ExecuteReader()

      If myReader.Read() Then
        If myReader("xPostDate") IsNot DBNull.Value Then
          xPostDateText.Text = Date.Parse(myReader("xPostDate")).ToShortDateString()
        End If
        QuestionText.Text = "<h3>標題：" & myReader("sTitle") & "</h3>"
        QuestionText.Text &= "<table class=""subtable"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""排版用表格"">"
        QuestionText.Text &= "<tr><td class=""subleft"">"
		if myReader("xBody") IsNot DBNull.Value Then
			QuestionText.Text &= "<div class=""problem"">" & myReader("xBody").replace(Chr(13), "<br />") & "</div>"
		else
			QuestionText.Text &= "<div class=""problem"">&nbsp;</div>"
		end if
        QuestionText.Text &= "</td></tr></table>"
        If myReader("xNewWindow") = "Y" Then
          DiscussFlag = False
        End If
      Else
        Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "linkerror", Script.Replace("{0}", "文章連結錯誤!!"))
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
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

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
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Try
      sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, KnowledgeForum.DiscussCount, "
      sqlString &= "KnowledgeForum.CommandCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, Member.realname, ISNULL(Member.nickname, '') AS nickname "
      sqlString &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
      sqlString &= "WHERE (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') "
      sqlString &= "AND (KnowledgeForum.Status = 'N') ORDER BY CuDTGeneric.xPostDate DESC"
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
      myReader = myCommand.ExecuteReader()

      myTable = New DataTable
      myTable.Load(myReader)

      Total = myTable.Rows.Count()
      sb = New StringBuilder

      sb.Append("<div class=""Dis""><h4>討論(" & totalDiscuss.ToString & ")</h4>")
      'sb.Append("<div class=""Dis""><h4><a href=""#"">討論</a></h4>")
      'sb.Append("<div class=""slidertable""><table border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""排版用表格"">")
      'sb.Append("<tr><td>摘要字數：(少)</td><td valign=""middle"" class=""slider""><img src=""images/btn_slider.gif"" alt=""sliderbar"" /></td>")
      'sb.Append("<td>(多)</td></tr></table></div>")

      If MemberId <> "" And MemberId IsNot Nothing And MemberActor = "005" Then

        sb.Append("<div class=""float2""><a href=""/knowledge/knowledge_professional_add.aspx?ArticleId=" & Request.QueryString("ArticleId") & _
                "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>專家補充</a></div>")

      End If
      If DiscussFlag Then
        sb.Append("<div class=""float2""><a href=""/knowledge/knowledge_discuss.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>參加討論</a></div>")
      Else
        sb.Append("<div class=""float2""></div>")
      End If

      DisTitle.Text = sb.ToString()
      sb = Nothing

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

        myDataRow = myTable.Rows.Item(Position)

        If PageNumber > 1 Then
          PreviousText.Visible = True
          PreviousImg.Visible = True
          PreviousLink.NavigateUrl = "/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Visible = False
          PreviousImg.Visible = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextText.Visible = True
          NextImg.Visible = True
          NextLink.NavigateUrl = "/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Visible = False
          NextImg.Visible = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""
        sb = New StringBuilder

        For i = 0 To PageSize - 1

          sb.Append("<div class=""DisList""><ul class=""Issued"">")

          If myDataRow("nickname") <> "" Then
            sb.Append("<li>" & myDataRow("nickname").ToString.Trim & "</li>")
          Else
            Dim name As String = myDataRow("realname").ToString.Trim
			if instr(name,"&#")>0 then
				name = name.Substring(0, 8) & "*" & name.Substring(name.Length - 8, 8)
			else
				name = name.Substring(0, 1) & "*" & name.Substring(name.Length - 1, 1)
			end if
            sb.Append("<li>" & name & "</li>")
          End If

          sb.Append("<li>發表於 ")
          If myDataRow("xPostDate") IsNot DBNull.Value Then
            sb.Append(Date.Parse(myDataRow("xPostDate")).ToShortDateString())
          End If
          sb.Append("</li>")
          sb.Append("<li>意見數：" & myDataRow("CommandCount") & "</li>")
          sb.Append(ConcatStarString("1", Integer.Parse(myDataRow("GradeCount")), Integer.Parse(myDataRow("GradePersonCount"))))
          sb.Append("</ul><p>")
          If myDataRow("xBody").ToString().Length > 200 Then
            sb.Append(myDataRow("xBody").ToString().Substring(0, 200).Replace("<", "&lt;").Replace(">", "&gt;").Replace(Chr(13), "<br />"))
          Else
            sb.Append(myDataRow("xBody").replace(Chr(13), "<br />"))
          End If

          sb.Append("...<a href=""/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & myDataRow("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>詳全文</a></p></div>")

          Position += 1
          If myTable.Rows.Count <= Position Then
            Exit For
          End If
          myDataRow = myTable.Rows(Position)
        Next

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
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

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
End Class
