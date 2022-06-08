Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Pedia_PediaContent
  Inherits System.Web.UI.Page

  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sql As String
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader
  Protected accountTag As String = ""
  Protected haveWordBody As String = "none"
  Protected haveWordKeyword As String = "none"
  Protected haveEngTitle As String = "none"
  
  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ArticleId As String = ""
    Dim canAdd As Boolean = False

    ArticleId = Request.QueryString("AId")

    If Not Page.IsPostBack Then
            '把count移到viewcounter.aspx 
            'UpdateClickCount(ArticleId)

      If Not FillWordContent(ArticleId, canAdd) Then
        ShowErrorMessage()
        Exit Sub
      End If

      '---kpi use---reflash wont add grade---
      If Request.QueryString("kpi") <> "0" Then
        '---start of kpi user---20080911---vincent---
        If Session("memID") IsNot Nothing Or Session("memID") <> "" Then
          Dim browse As New KPIBrowse(Session("memID"), "browsePediaWordCP", ArticleId)
          browse.HandleBrowse()
        End If
        '---end of kpi use---
        Dim relink As String = "/Pedia/PediaContent.aspx?AId=" & ArticleId & "&kpi=0"
        Response.Redirect(relink)
        Response.End()
      End If
      '---end of kpi use---   

      If canAdd Then
        CanAddLink.Visible = True
        CanAddLink.NavigateUrl = "/Pedia/PediaExplainAdd.aspx?AId=&PAId=" & ArticleId
      End If

      ExplainListLink.Visible = True
      ExplainListLink.NavigateUrl = "/Pedia/PediaExplainList.aspx?PAId=" & ArticleId

	  If Not FillExplainContent(ArticleId) Then
		ShowErrorMessage()
		Exit Sub
	  End If

    End If

  End Sub

  Sub UpdateClickCount(ByVal id As String)

    myConnection = New SqlConnection(ConnString)
    Try
      sql = " UPDATE CuDTGeneric SET ClickCount = ClickCount + 1 WHERE (CuDTGeneric.iCUItem = @icuitem)"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", id)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "詞目錯誤"      
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  Function FillWordContent(ByVal id As String, ByRef canAdd As Boolean) As Boolean

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT CuDTGeneric.sTitle, CuDTGeneric.xBody, CuDTGeneric.xKeyword, CuDTGeneric.vGroup, Pedia.engTitle, "
      sql &= " Pedia.formalName, Pedia.localName FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (CuDTGeneric.iCUItem = @icuitem)"
            'response.write (sql)
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", id)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then

        WordTitle.Text = "<a href=""javascript:WordSearch('" & myReader("sTitle") & "')"">" & myReader("sTitle") & "</a>"
        if Not String.IsNullOrEmpty(myReader("engTitle"))
			haveEngTitle = ""
		End if
		WordEngTitle.Text = IIf(myReader("engTitle") Is DBNull.Value Or myReader("engTitle") = "", "&nbsp;", myReader("engTitle"))
        WordFormalName.Text = IIf(myReader("formalName") Is DBNull.Value Or myReader("formalName") = "", "&nbsp;", myReader("formalName"))
        WordLocalName.Text = IIf(myReader("localName") Is DBNull.Value Or myReader("localName") = "", "&nbsp;", myReader("localName"))
        WordBody.Text = myReader("xBody").ToString.Replace(vbCrLf, "<br />")
		If Not String.IsNullOrEmpty(myReader("xBody"))
			haveWordBody = ""
		End If
        If myReader("xKeyword") Is DBNull.Value Or myReader("xKeyword") = "" Then
          WordKeyword.Text = "&nbsp;"
        Else
          Dim items = myReader("xKeyword").ToString.Split(";")
		  haveWordKeyword = ""
          For Each item As String In items
            If item IsNot Nothing And item <> "" Then
              WordKeyword.Text &= "<a href=""javascript:WordSearch('" & item & "')"">" & item & "</a>;"
            End If
          Next
        End If

        If myReader("vGroup") IsNot DBNull.Value Then
          If myReader("vGroup").ToString = "Y" Then
            canAdd = True
          End If
        End If

      Else
        Session("PediaErrMsg") = "詞目不存在"
        Return False
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "詞目錯誤"
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return True

  End Function

  Function FillExplainContent(ByVal id As String) As Boolean

    Dim sb As New StringBuilder
    Dim index As Integer = 1

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT TOP (3) CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xBody, Pedia.parentIcuItem, Member.realname, Member.nickname, Member.account, "
      sql &= " Pedia.commendTime FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem INNER JOIN Member ON Pedia.memberId = Member.account "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (Pedia.parentIcuitem = @icuitem) ORDER BY Pedia.commendTime DESC"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", id)
      myReader = myCommand.ExecuteReader

      sb.Append("<table class=""type02"" summary=""結果條列式"">")
      sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th><th scope=""col"">補充解釋(摘要)</th><th scope=""col"" width=""10%"">發表者</th><th scope=""col"" width=""10%"">等級</th><th scope=""col"" width=""10%"">發佈日期</th></tr>")

      While myReader.Read

        sb.Append("<tr><td>" & index & "</td>")
        sb.Append("<td><a href=""/Pedia/PediaExplainContent.aspx?AId=" & myReader("iCUItem") & "&PAId=" & myReader("parentIcuItem") & """>")
        If myReader("xBody").ToString.Length > 70 Then
          sb.Append(NoHTML(myReader("xBody").ToString.Substring(0, 70)))
        Else
          sb.Append(NoHTML(myReader("xBody").ToString))
        End If
        sb.Append("</a></td>")
		
		Dim showName As String
        If myReader("nickname") IsNot DBNull.Value Then
          If myReader("nickname") <> "" Then
            'sb.Append("<td>" & myReader("nickname").ToString & "</td>")
			showName = myReader("nickname").ToString
          Else
            'sb.Append("<td>" & myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1) & "</td>")
			showName = myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1)
          End If
        ElseIf myReader("realname") IsNot DBNull.Value Then          
          'sb.Append("<td>" & myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1) & "</td>")
		  showName = myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1)
        Else
          'sb.Append("<td>&nbsp;</td>")
		  showName = "&nbsp;"
        End If
		
		'===========在補充解釋清單中, 顯示每個[發表者]的等級 2009.08.18 by ivy

			Dim account As String = myReader("account").ToString.Trim
			Dim gradeLevel As String = GetMemberGradeLevel(account)
			Dim gradeDesc As String =""
			
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
			
			'JIRA COAKM-36 討論中可以看到參予討論人員(所有會員，不限制是達人)的平均評價
			sb.Append("<td>" & MemberInfo(account,showName,Convert.ToInt32(gradeLevel)) & "</td>")
			
			sb.Append("<td>" & gradeDesc & "</td>")			
	    '===========End by ivy
        sb.Append("<td>" & Date.Parse(myReader("commendTime")).ToShortDateString & "</td>")
        sb.Append("</tr>")

        index += 1

      End While

      sb.Append("</table>")

      ExplainText.Text = sb.ToString

      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "補充解釋錯誤"
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return True

  End Function
  
  Function MemberInfo(ByVal account As String,ByVal name As String, ByVal gradeLevel As Integer) As String
	
	Dim string1
	string1 = MemberInfoUtility.ReplaceMemberName(account,name,gradeLevel)
	account = string1(0)
	accountTag &= string1(1)

    return account
  
  End Function
  

  '取得會員的知識等級
  Function GetMemberGradeLevel(ByVal account As String) As String
	Using conn As New SqlConnection(ConnString)
		conn.Open()
		Dim sqlMemLevel As String =""
		Dim gradeLevel As String =""
		Dim memLevelCmd As SqlCommand
		Dim memLevelReader As SqlDataReader
		sqlMemLevel = "SELECT top 1 mCode FROM CodeMain WHERE (codeMetaID = 'gradelevel') "
		sqlMemLevel & = "AND ((SELECT calculateTotal FROM MemberGradeSummary WHERE memberId = @account)>= mValue) ORDER BY mSortValue DESC"
		memLevelCmd = New SqlCommand(sqlMemLevel, conn)
		memLevelCmd.Parameters.AddWithValue("@account", account)
		gradeLevel = memLevelCmd.ExecuteScalar()
		
		
			
		Return gradeLevel
	End using
  End Function

  Function NoHTML(ByVal Htmlstring As String) As String

    Htmlstring = Regex.Replace(Htmlstring, "<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "<(.[^>]*)>", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "([\r\n])[\s]+", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "-->", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "<!--.*", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(quot|#34);", "\", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(amp|#38);", "&", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(lt|#60);", "<", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(gt|#62);", ">", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(nbsp|#160);", "   ", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(cent|#162);", "\xa2", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(pound|#163);", "\xa3", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(copy|#169);", "\xa9", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&#(\d+);", "", RegexOptions.IgnoreCase)
    Htmlstring = Htmlstring.Replace("<", "")
    Htmlstring = Htmlstring.Replace(">", "")
    Htmlstring = Htmlstring.Replace("\r\n", "")
    'Htmlstring = HttpContext.Current.Server.HtmlEncode(Htmlstring).Trim()
    Return Htmlstring

  End Function

  Sub ShowErrorMessage()

    Dim str = "<script>alert('" & Session("PediaErrMsg") & "');window.location.href='/Pedia/PediaList.aspx';</script>"

    Session("PediaErrMsg") = ""

    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "PediaError", str)

  End Sub

End Class
