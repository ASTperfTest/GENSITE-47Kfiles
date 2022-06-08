Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_knowledge_opinion_lp
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ArticleId As String
    Dim DArticleId As String
    Dim ArticleType As String
    Dim CategoryId As String
    Dim PageNumber As Integer
    Dim PageSize As Integer

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

    If Request.QueryString("DArticleId") = "" Then
      Response.Redirect("/knowledge/knowledge.aspx")
      Response.End()
    Else
      DArticleId = Request.QueryString("DArticleId")
    End If

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
        QuestionText.Text &= "<div class=""problem"">" & myReader("xBody").replace(Chr(13), "<br />") & "</div>"
        QuestionText.Text &= "</td></tr></table>"
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

    Try
      sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, KnowledgeForum.DiscussCount, "
      sqlString &= "KnowledgeForum.CommandCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, Member.realname, ISNULL(Member.nickname, '') AS nickname "
      sqlString &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account "
      sqlString &= "WHERE (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') "
      sqlString &= "AND (KnowledgeForum.Status = 'N') ORDER BY CuDTGeneric.xPostDate DESC"

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", DArticleId)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeOpinionCtUnitId"))
      myReader = myCommand.ExecuteReader()

      myTable = New DataTable
      myTable.Load(myReader)

      Total = myTable.Rows.Count()
      sb = New StringBuilder

      sb.Append("<div class=""Dis""><h4>意見</h4>")

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
          PreviousLink.NavigateUrl = "/knowledge/knowledge_opinion_lp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Visible = False
          PreviousImg.Visible = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextText.Visible = True
          NextImg.Visible = True
          NextLink.NavigateUrl = "/knowledge/knowledge_opinion_lp.aspx?ArticleId=" & ArticleId & "&DArticleId=" & DArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
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
          sb.Append("<li></li>")
		  'added by Joey,2009/09/21，問題單回應，若已檢舉過該意見，且站方尚未處理完畢，則不能再次檢舉
		  dim sqlStatement as string=String.empty
		  Dim myCommand2 As SqlCommand
		  Dim myReader2 As SqlDataReader    
		  Dim myTable2 As DataTable    
		  
		  sqlStatement = " select * from KNOWLEDGE_REPORT "
		  sqlStatement &= " where STATUS=0 AND ARTICLE_ID=@icuitem AND CREATOR=@memID"
		  

		  'myConnection.Open()
		  myCommand2 = New SqlCommand(sqlStatement, myConnection)
		  myCommand2.Parameters.AddWithValue("@icuitem", myDataRow("iCUItem"))
		  myCommand2.Parameters.AddWithValue("@memID", Session("memID"))
		  myReader2 = myCommand2.ExecuteReader()

		  myTable2 = New DataTable
		  myTable2.Load(myReader2)
		  If myTable2.Rows.Count = 0 Then
			sb.Append("<li><a href=""#"" onclick=""window.showModalDialog('http://" & Request.Url.Authority & "/mailbox.asp?a=a9191',self);return false;"">檢舉</a></li>")
		  Else
			sb.Append("<li>(站方正在處理您檢舉的意見...)</li>")
		  End If	
		    
		  sb.Append("<input type=""hidden""  name=""type"" value=""3"">" )
		  sb.Append("<li><input type=""hidden""  name=""ARTICLE_ID"" value=" & myDataRow("iCUItem") & "></li>")		  
		  
          sb.Append("</ul><p>")

          sb.Append(myDataRow("xBody").ToString().Replace("<", "&lt;").Replace(">", "&gt;").Replace(Chr(13), "<br />"))

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


End Class
