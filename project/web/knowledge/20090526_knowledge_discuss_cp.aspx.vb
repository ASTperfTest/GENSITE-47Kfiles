Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_discuss_cp
    Inherits System.Web.UI.Page


  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim DArticleId As String = ""
    Dim ArticleId As String = ""
    Dim ArticleType As String = ""
    Dim CategoryId As String = ""
	Dim Script1 As String = ""
	'type =1 是發表文章 這時才要檢查有沒有會員
	If Session("memID") = "" Or Session("memID") Is Nothing Then
		if request("type") = "1" then
			Script1 = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge_cp.aspx?ArticleId="&request("ArticleId")&"&ArticleType="&request("ArticleType")&"&CategoryId="&request("CategoryId")&"&kpi="&request("kpi")&"';</script>"
			Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script1)
			Exit Sub
		end if
	end if

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

    If Request.QueryString("DArticleId") = "" Then
      Response.Redirect("/knowledge/knowledge.aspx")
      Response.End()
    Else
      Try
        DArticleId = Integer.Parse(Request.QueryString("DArticleId"))
      Catch ex As Exception
        Response.Redirect("/knowledge/knowledge.aspx")
        Response.End()
      End Try     
    End If

    '---kpi use---reflash wont add grade---
    If Request.QueryString("kpi") <> "0" Then
      '---start of kpi user---20080911---vincent---
      If Session("memID") IsNot Nothing Or Session("memID") <> "" Then
        Dim browse As New KPIBrowse(Session("memID"), "browseDiscussCP", DArticleId)
        browse.HandleBrowse()
      End If
      '---end of kpi use---
      Dim relink As String = "/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&kpi=0"
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
	'BackText.Text = "<a href=""/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """ class=""Back"">回上一頁</a>"

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

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()

      '---更新瀏覽數---
      'sqlString = "UPDATE KnowledgeForum SET BrowseCount = BrowseCount + 1 WHERE gicuitem = @dgicuitem"
      'myCommand = New SqlCommand(sqlString, myConnection)            
      'myCommand.Parameters.AddWithValue("@dgicuitem", DArticleId)
      'myCommand.ExecuteNonQuery()
      'myCommand.Dispose()

      sqlString = "SELECT CuDTGeneric.sTitle, ISNULL(CuDTGeneric.xBody, '') AS xBody, CuDTGeneric.xPostDate FROM KnowledgeForum INNER JOIN CuDTGeneric ON "
      sqlString &= "KnowledgeForum.gicuitem = CuDTGeneric.iCUItem WHERE (CuDTGeneric.iCUItem = @ArticleId) AND (CuDTGeneric.fCTUPublic = 'Y') AND (KnowledgeForum.Status = 'N')"

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
        QuestionText.Text &= "<div class=""problem"">" & myReader("xBody").replace(Chr(13), "<br />") & "&nbsp;</div>"
        QuestionText.Text &= "</td></tr></table>"
      Else
        Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "linkerror", Script.Replace("{0}", "文章連結錯誤!!"))
        Exit Sub
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

      '---抓出內容---
      sqlString = "SELECT CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount, KnowledgeForum.GradeCount, "
      sqlString &= "KnowledgeForum.GradePersonCount, Member.realname, ISNULL(Member.nickname, '') AS nickname FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sqlString &= "INNER JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCUItem = @ArticleId) AND (CuDTGeneric.fCTUPublic = 'Y') AND (KnowledgeForum.Status = 'N')"

      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ArticleId", DArticleId)
      myReader = myCommand.ExecuteReader()

      If myReader.Read() Then

        sb = New StringBuilder
        sb.Append("<div class=""DisContent"">")
        sb.Append("<div class=""DisText"">")
        sb.Append("<ul class=""Issued"">")

        If myReader("nickname") <> "" Then
          sb.Append("<li>" & myReader("nickname").ToString.Trim & "</li>")
        Else
          Dim name As String = myReader("realname").ToString.Trim
		  if instr(name,"&#")>0 then
			name = name.Substring(0, 8) & "*" & name.Substring(name.Length - 8, 8)
		  else
			name = name.Substring(0, 1) & "*" & name.Substring(name.Length - 1, 1)
		  end if
          sb.Append("<li>" & name & "</li>")
        End If

        sb.Append("<li>發表於")
        If myReader("xPostDate") IsNot DBNull.Value Then
          sb.Append(Date.Parse(myReader("xPostDate")).ToShortDateString)
        End If
        sb.Append("</li>")
        sb.Append("<li>意見數：" & myReader("CommandCount") & "</li>")
        sb.Append(ConcatStarString(Integer.Parse(myReader("GradeCount")), Integer.Parse(myReader("GradePersonCount"))))
        sb.Append("</ul>")
                sb.Append("<p>" & myReader("xBody").replace(Chr(13), "<br />") & "</p>")

        ArticleText.Text = sb.ToString()
        sb = Nothing

      Else
        Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
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
      Response.Redirect("/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    ArticleText.Text &= GetPicInfo(DArticleId)
    ArticleText.Text &= "</div></div>"

    '---意見列表---
    Try
      sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, KnowledgeForum.DiscussCount, "
      sqlString &= "KnowledgeForum.CommandCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, Member.realname, ISNULL(Member.nickname, '') AS nickname "
      sqlString &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
      sqlString &= "WHERE (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.iCTUnit = @ictunit) AND (KnowledgeForum.Status = 'N') ORDER BY CuDTGeneric.xPostDate DESC"
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", DArticleId)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeOpinionCtUnitId"))
      myReader = myCommand.ExecuteReader()

      myTable = New DataTable
      myTable.Load(myReader)

      Total = myTable.Rows.Count()
      Dim i As Integer = 0
      sb = New StringBuilder

      sb.Append("<div class=""Dis""><h4>意見(" & Total.ToString & ")</h4>")

      If Total <> 0 Then

        For i = 0 To 2
          myDataRow = myTable.Rows(i)
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
          sb.Append("<p>" & myDataRow("xBody").replace(Chr(13), "<br />") & "</p></div>")
          If myTable.Rows.Count = i + 1 Then
            Exit For
          End If
        Next
        sb.Append("<div class=""float""><a href=""/knowledge/knowledge_opinion_lp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>看全部意見</a></div></div>")
      End If

      DiscussText.Text = sb.ToString
      sb = Nothing

      myCommand.Dispose()

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Response.Redirect("/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    '---回應區---
    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      PostCommentTBox.Visible = False
      StarValue.Visible = False
      ConfirmBtn.Visible = False
      CancelBtn.Visible = False
      stararea.Visible = False
      CaptChaText1.Visible = False
      CaptChaText2.Visible = False
      MemberCaptChaImage.Visible = False
      MemberCaptChaTBox.Visible = False
      PostCommentLabel.Text = "歡迎大家參予回應，此功能僅開放會員進行回應與張貼，如您已是會員，請登入後再行回應，如未註冊任何會員，歡迎您先申請成為會員"

    Else
      PostCommentLabel.Text = "發表意見："
    End If
    '-----------------------------------------------------------------------------
    Star1ImgLink.NavigateUrl = "~/knowledge/knowledge_grade_add.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Grade=1"
    Star2ImgLink.NavigateUrl = "~/knowledge/knowledge_grade_add.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Grade=2"
    Star3ImgLink.NavigateUrl = "~/knowledge/knowledge_grade_add.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Grade=3"
    Star4ImgLink.NavigateUrl = "~/knowledge/knowledge_grade_add.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Grade=4"
    Star5ImgLink.NavigateUrl = "~/knowledge/knowledge_grade_add.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Grade=5"

  End Sub

  Function ConcatStarString(ByVal GradeCount As Integer, ByVal PersonCount As Integer) As String

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
    str &= "(" & PersonCount & "人評價)</li>"

    Return str

  End Function

  Protected Sub ConfirmBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ConfirmBtn.Click

    Dim PicInput As Boolean = False
    '---handle picture---
    If (Session("CaptchaImageText") = Nothing) Then
      ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼失效!!');window.location.href='" & Request.RawUrl() & "';</script>")
      Exit Sub
    ElseIf (MemberCaptChaTBox.Text <> Session("CaptchaImageText").ToString()) Then
      ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼錯誤!!');history.back();</script>")
      Exit Sub
    End If

    Dim MemberId As String
    Dim Script As String = "<script>alert('{0}!!');history.back();</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    Dim ArticleType As String
    Dim CategoryId As String

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

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As New SqlConnection(ConnString)
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim ArticleId As String = Request.QueryString("ArticleId")
    Dim DArticleId As String = Request.QueryString("DArticleId")
    Dim MaxIcuitem As String
    Dim Flag As Boolean = False
    Dim ArticleContent As String = PostCommentTBox.Text.Trim
    ArticleContent = ArticleContent.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")

    '---check if the member has post opinion---
    Try
      sqlString = "SELECT CuDTGeneric.iCUItem FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.iEditor = @ieditor) AND (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N')"
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeOpinionCtUnitId"))
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)
      myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
      myCommand.Parameters.AddWithValue("@icuitem", DArticleId)
      myReader = myCommand.ExecuteReader
      If myReader.HasRows Then
        Flag = True
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
      Response.Redirect("/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    If Flag Then
      Dim ScriptHaveOpinion As String = "<script>alert('您已對此篇文章給予意見!!');history.back();</script>"
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "submit", ScriptHaveOpinion)
      Exit Sub
    End If

    Try
      '---insert discuss opinion---
      sqlString = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, xBody, siteId) "
      sqlString &= "VALUES (@ibasedsd, @ictunit, 'Y', @stitle, @ieditor, '0', @xbody, @siteid)"
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ibasedsd", WebConfigurationManager.AppSettings("KnowledgeBaseDSD"))
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeOpinionCtUnitId"))
      myCommand.Parameters.AddWithValue("@stitle", "討論意見-" & DArticleId)
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)
      myCommand.Parameters.AddWithValue("@xbody", ArticleContent)
      myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sqlString = "SELECT @@identity AS MaxIcuitem"
      Dim MaxCommand As SqlDataAdapter = New SqlDataAdapter(sqlString, myConnection)
      Dim MaxDataSet As DataSet = New DataSet()
      MaxCommand.Fill(MaxDataSet, "IndexRecommend")
      Dim myTable As DataTable = MaxDataSet.Tables("IndexRecommend")
      MaxIcuitem = myTable.Rows(0).Item(0)

      sqlString = "INSERT INTO KnowledgeForum(gicuitem, ParentIcuitem) VALUES (@gicuitem, @parenticuitem)"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@gicuitem", MaxIcuitem)
      myCommand.Parameters.AddWithValue("@parenticuitem", DArticleId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      '---update the parent report---
      sqlString = "UPDATE KnowledgeForum SET CommandCount = CommandCount + 1 WHERE gicuitem = @gicuitem or gicuitem = @agicuitem"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@agicuitem", ArticleId)
      myCommand.Parameters.AddWithValue("@gicuitem", DArticleId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      '---20080927---vincent---kpi share---
      Dim share As New KPIShare(MemberId, "shareOpinion", MaxIcuitem)
      share.HandleShare()

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Response.Redirect("/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Dim SuccessScript As String = "<script>alert('發表意見成功');window.location.href='/knowledge/knowledge_discuss_cp.aspx?ArticleId=" & ArticleId & _
              "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "submit", SuccessScript)

  End Sub

  Function GetPicInfo(ByVal aid As String) As String

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
        str &= "<div class=""knowledgepic"">"
        str &= "<img src=""" & myReader("picPath") & """ alt=""" & myReader("picTitle") & """ />"
        str &= "<p>" & myReader("picTitle") & "</p>"
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
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return str

  End Function

  Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click

    Dim ArticleType As String
    Dim CategoryId As String

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

    Response.Redirect("/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & Request.QueryString("ArticleId") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)

  End Sub

End Class
