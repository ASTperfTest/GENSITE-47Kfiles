Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_professional
    Inherits System.Web.UI.Page

  Dim ArticleId As String
  Dim DArticleId As String
  Dim ArticleType As String
  Dim CategoryId As String
  Dim MemberId As String
  Dim MemberActor As String

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      MemberActor = ""
    Else
      MemberId = Session("memID").ToString()
      MemberActor = Session("gstyle").ToString()
    End If

    '-----------------------------------------------------------------------------

    If Request.QueryString("ArticleType") <> "A" And Request.QueryString("ArticleType") <> "B" And Request.QueryString("ArticleType") <> "C" _
        And Request.QueryString("ArticleType") <> "D" And Request.QueryString("ArticleType") <> "E" Then
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

    If Request.QueryString("ArticleId") = "" Then
      Response.Redirect("/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
      Response.End()
    Else
      ArticleId = Request.QueryString("ArticleId")
    End If

    If Request.QueryString("DArticleId") = "" Then
      Response.Redirect("/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
      Response.End()
    Else
      DArticleId = Request.QueryString("DArticleId")
    End If

    PathText.Text = "&gt;<a href=""/knowledge/knoledge_lp.aspx"">知識分類</a>"

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

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim sb As New StringBuilder
    Dim HavePros As String = ""

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "SELECT KnowledgeForum.HavePros, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.iEditor FROM KnowledgeForum INNER JOIN CuDTGeneric ON "
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

        HavePros = myReader("HavePros")
        sb = Nothing

      Else
        Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
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

    If HavePros = "Y" Then
      sb = New StringBuilder
      Try
        sqlString = "SELECT Member.realname, ISNULL(Member.nickname, '') AS nickname, CuDTGeneric.xBody, CuDTGeneric.xPostDate, CuDTGeneric.iEditor FROM CuDTGeneric INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
        sqlString &= "INNER JOIN KnowledgeForum ON KnowledgeForum.gicuitem = CuDTGeneric.iCUItem "
        sqlString &= "WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCUItem = @icuitem) "
        sqlString &= "AND (CuDTGeneric.iCTUnit = @ictunit) ORDER BY CuDTGeneric.xPostDate DESC "
        myConnection.Open()
        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@icuitem", DArticleId)
        myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
        myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("ProfessionAnswerCtUnitId"))
        myReader = myCommand.ExecuteReader()
        If myReader.Read() Then

          sb.Append("<div class=""experts""><h4>駐站專家補充</h4>")
          sb.Append("<div class=""exphoto""><div class=""phoimg""><img src=""images/expert.gif"" alt=""圖片說明"" /></div>")
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
          sb.Append("<p>" & myReader("xBody").replace(Chr(13), "<br />") & "</p></div>")

          If myReader("iEditor") = MemberId Then
            DeleteBtn.Visible = True
          End If
        Else
          Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
          Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "linkerror", Script.Replace("{0}", "文章連結錯誤!!"))
          Exit Sub
        End If
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

  End Sub

    Protected Sub DeleteBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteBtn.Click

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim Flag As Boolean = False
  
        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT CuDTGeneric.iEditor FROM KnowledgeForum INNER JOIN CuDTGeneric ON "
            sqlString &= "KnowledgeForum.gicuitem = CuDTGeneric.iCUItem WHERE (CuDTGeneric.iCUItem = @ArticleId) AND (CuDTGeneric.fCTUPublic = 'Y') AND (KnowledgeForum.Status = 'N') "
            sqlString &= "AND (CuDTGeneric.iEditor = @ieditor)"
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@ArticleId", DArticleId)
            myCommand.Parameters.AddWithValue("@ieditor", MemberId)
            myReader = myCommand.ExecuteReader()

            If myReader.Read() Then

                Flag = True
            Else
                Flag = False
                Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "linkerror", Script.Replace("{0}", "文章連結錯誤!!"))
                Exit Sub
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            If Flag Then

                sqlString = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem = @gicuitem"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@gicuitem", DArticleId)                
                myCommand.ExecuteNonQuery()
                myCommand.Dispose()
                Dim Script As String = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "linkerror", Script.Replace("{0}", "專家補充刪除成功!!"))
                Exit Sub
            End If

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

End Class
