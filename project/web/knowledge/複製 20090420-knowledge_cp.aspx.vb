Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_cp
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ArticleId As String = ""
    Dim ArticleType As String = ""
    Dim CategoryId As String = ""
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
      ArticleId = Request.QueryString("ArticleId")
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
        Dim browse As New KPIBrowse(Session("memID"), "browseQuestionCP", ArticleId)
        browse.HandleBrowse()
      End If
      '---end of kpi use---
      Dim relink As String = "/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&kpi=0"
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
    BackText.Text = "<a href=""/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """ class=""Back"">回上一頁</a>"

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
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try
	'計算討論則數 end
	
    Try
      myConnection.Open()

      '---更新瀏覽數---
      sqlString = "UPDATE KnowledgeForum SET BrowseCount = BrowseCount + 1 WHERE gicuitem = @agicuitem"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@agicuitem", ArticleId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sqlString = "UPDATE CuDTGeneric SET ClickCount = (SELECT BrowseCount FROM KnowledgeForum WHERE gicuitem = @agicuitem) WHERE icuitem = @agicuitem"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@agicuitem", ArticleId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sqlString = "SELECT KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount, KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, "
      sqlString &= "KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount,  CuDTGeneric.xNewWindow, ISNULL(CuDTGeneric.vGroup, '') AS vGroup, "
      sqlString &= "KnowledgeForum.HavePros, KnowledgeForum.ParentIcuitem, CuDTGeneric.iCUItem, CuDTGeneric.iCTUnit, CuDTGeneric.sTitle, CuDTGeneric.topCat, "
      sqlString &= "CuDTGeneric.xKeyword, CuDTGeneric.xPostDate, ISNULL(CuDTGeneric.xBody, '') AS xBody, Member.realname, ISNULL(Member.nickname, '') AS nickname FROM KnowledgeForum INNER JOIN CuDTGeneric ON "
      sqlString &= "KnowledgeForum.gicuitem = CuDTGeneric.iCUItem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCUItem = @ArticleId) "
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
        sb.Append("<div class=""problem"">" & myReader("xBody").replace(Chr(13), "<br />") & "&nbsp;</div>")

        sb.Append(GenerateQuestionAdditional(myReader("iCUItem")))

        sb.Append(GetPicInfo(myReader("iCUItem")))


        sb.Append("<td class=""subright""><div><ul>")

        If ActivityFlag Then
          If myReader("vGroup") = "A" Then
            sb.Append("<li>本問題符合知識一起答活動範圍")
            sb.Append("<div class=""staricon"">")
                        sb.Append("<img class=rating src=""../xslGip/style3/images/good.gif"" align=top>")
            sb.Append("</div></li><br/><br/>")
          End If
        End If

        sb.Append("<li>發問者：")
        If myReader("nickname") <> "" Then
          sb.Append(myReader("nickname").ToString.Trim)
        Else
          Dim name As String = myReader("realname").ToString.Trim
		  if instr(name,"&#")>0 then
			name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
		  else
			name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
		  end if
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
		sb.Append("<li>瀏覽數：" & myReader("BrowseCount") & "</li><li>討論數：" & totalDiscuss.ToString  & "</li>")
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
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    If HavePros = "Y" Then
      sb = New StringBuilder
      Try
        sqlString = "SELECT Member.realname, ISNULL(Member.nickname, '') AS nickname, ISNULL(Member.photo, '') AS photo, CuDTGeneric.xBody, CuDTGeneric.xPostDate, CuDTGeneric.iCUItem "
        sqlString &= "FROM CuDTGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account INNER JOIN KnowledgeForum ON KnowledgeForum.gicuitem = CuDTGeneric.iCUItem "
        sqlString &= "WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N') AND (KnowledgeForum.ParentIcuitem = @icuitem) "
        sqlString &= "AND (CuDTGeneric.iCTUnit = @ictunit) ORDER BY CuDTGeneric.xPostDate DESC "
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
            sb.Append("<img src=""/Public/Data/" & myReader("photo") & """ alt=""圖片說明"" />")
          End If
          sb.Append("</div>")
          If myReader("nickname") <> "" Then
            sb.Append("<p>" & myReader("nickname").ToString.Trim)
          Else
            Dim name As String = myReader("realname").ToString.Trim
			if instr(name,"&#")>0 then
				name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
			else
				name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
			end if
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
          If myReader("xBody").ToString().Length > 200 Then
            sb.Append("<p>" & myReader("xBody").ToString().Substring(0, 200).Replace(Chr(13), "<br />"))
          Else
            sb.Append("<p>" & myReader("xBody").replace(Chr(13), "<br />"))
          End If
          sb.Append("...<a href=""/knowledge/knowledge_professional.aspx?ArticleId=" & ArticleId & "&DArticleId=" & myReader("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>詳全文</a></p></div>")

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
        If myConnection.State = ConnectionState.Open Then
          myConnection.Close()
        End If
      End Try
    Else
      ProsText.Text = ""
    End If




    Try
      sqlString = "SELECT TOP 3 CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, KnowledgeForum.DiscussCount, "
      sqlString &= "KnowledgeForum.CommandCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, Member.realname, ISNULL(Member.nickname, '') AS nickname "
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

      If MemberId <> "" And MemberId IsNot Nothing And MemberActor = "005" Then

        sb.Append("<div class=""float2""><a href=""/knowledge/knowledge_professional_add.aspx?ArticleId=" & Request.QueryString("ArticleId") & _
                "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>專家補充</a></div>")

      End If
      If DiscussFlag Then
        sb.Append("<div class=""float2""><a href=""/knowledge/knowledge_discuss.aspx?ArticleId=" & Request.QueryString("ArticleId") & _
                    "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>參加討論</a></div>")
      Else
        sb.Append("<div class=""float2""></div>")
      End If


      If Total <> 0 Then

        For i = 0 To 3
          myDataRow = myTable.Rows.Item(i)
          sb.Append("<div class=""DisList""><ul class=""Issued"">")
          If myDataRow("nickname") <> "" Then
            sb.Append("<li>" & myDataRow("nickname").ToString.Trim & "</li>")
          Else
            Dim name As String = myDataRow("realname").ToString.Trim
			if instr(name,"&#")>0 then
				name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
			else
				name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
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
		  'sb.Append("</p></div>")

          If myTable.Rows.Count = i + 1 Then
            Exit For
          End If
        Next
        sb.Append("<div class=""float""><a href=""/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """>看全部討論</a></div></div>")
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
        str &= "<div class=""problem2"">" & myReader("xBody").replace(Chr(13), "<br />") & "</div>"
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
