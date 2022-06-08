Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class myknowledge_question_draft
    Inherits System.Web.UI.Page

  Dim ArticleId As String
  Dim MemberId As String
    Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>"
  Dim Script1 As String = "<script>alert('連結錯誤!!');window.location.href='/knowledge/myknowledge_question_lp.aspx';</script>"
  Dim ScriptFail As String = "<script>alert('系統錯誤!!');window.location.href='/knowledge/myknowledge_question_lp.aspx';</script>"
  Dim sb As StringBuilder
  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sqlString As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    If Request.QueryString("ArticleId") <> "" Then
      ArticleId = Request.QueryString("ArticleId")
    Else
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoArticleId", Script1)
      Exit Sub
    End If

    Dim ArticleType As String = Request.QueryString("ArticleType")
    Dim CategoryId As String = Request.QueryString("CategoryId")
    Dim Keyword As String = Web.HttpUtility.UrlDecode(Request.QueryString("Keyword"))
    Dim Sort As String = Request.QueryString("Sort")
    GoBackLink.NavigateUrl = "/knowledge/myknowledge_question_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort

    If Not Page.IsPostBack Then

      myConnection = New SqlConnection(ConnString)
      Try
        sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xBody, CuDTGeneric.xPostDate, CodeMain.mValue, KnowledgeForum.DiscussCount,ISNULL(CuDTGeneric.vGroup,'') AS vGroup, "
        sqlString &= "KnowledgeForum.CommandCount, KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, KnowledgeForum.GradeCount, "
        sqlString &= "KnowledgeForum.GradePersonCount FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode "
        sqlString &= "WHERE (CuDTGeneric.iCUItem = @icuitem) AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.iEditor = @memberid) AND (KnowledgeForum.Status = 'N') "

        myConnection.Open()

        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
        myCommand.Parameters.AddWithValue("@memberid", MemberId)
        myReader = myCommand.ExecuteReader()

        If myReader.Read Then
          QuestionTitleText.Text = myReader("sTitle")
          QuestionPostDateText.Text = Date.Parse(myReader("xPostDate")).ToShortDateString
          QuestionCategoryText.Text = myReader("mValue")
          QuestionBrowseCountText.Text = myReader("BrowseCount")
          QuestionDiscussCountText.Text = myReader("DiscussCount")
          QuestionTraceCountText.Text = myReader("TraceCount")
          QuestionCommandCountText.Text = myReader("CommandCount")
          AverageGradeCountText.Text = ConcatStarString(myReader("GradeCount"), myReader("GradePersonCount"))
          QuestionAdditionalContentText.Text = myReader("xBody") '.replace(chr(13), "<br />")
		  
			'若是活動期間,判斷是否為活動題目
			Dim Activity As New ActivityFilter(WebConfigurationManager.AppSettings("ActivityId"), MemberId)
			Dim ActivityFlag As Boolean = Activity.IsActivityEnd

			If ActivityFlag = False Then
				If myReader("vGroup") = "A" or myReader("vGroup") = "OFF" Then 
					QuestionConfirmBtn.visible = False
					QuestionSaveBtn.visible = False
					QuestionDelBtn.visible = False
				End If
			End If
			
			
		  
        Else
          Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
          Exit Sub
        End If

        If Not myReader.IsClosed Then
          myReader.Close()
        End If

        myCommand.Dispose()

        sqlString = "SELECT CuDTGeneric.xBody, CuDTGeneric.xPostDate FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
        sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.fCTUPublic = 'Y') AND (KnowledgeForum.Status = 'N')"

        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeAdditionalCtUnitId"))
        myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
        myReader = myCommand.ExecuteReader()

        If myReader.HasRows Then
          sb = New StringBuilder
          While myReader.Read
            sb.Append("<div class=""adddate3"">補充說明發表日期：" & Date.Parse(myReader("xPostDate")).ToShortDateString & "</div>")
            sb.Append("<div class=""problem3"">" & myReader("xBody").replace(Chr(13), "<br />") & "</div>")
          End While

          sb = Nothing
        Else

        End If

        If Not myReader.IsClosed Then
          myReader.Close()
        End If

        myCommand.Dispose()

      Catch ex As Exception
        If Request("debug") = "true" Then
          Response.Write(ex.ToString())
          Response.End()
        End If
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Fail", ScriptFail)
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

    Dim str As String = "<div class=""staricon"">"
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

    str &= "</div>(" & PersonCount & "人評價)"

    Return str

  End Function

  Protected Sub QuestionConfirmBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QuestionConfirmBtn.Click

    Dim ArticleType As String = Request.QueryString("ArticleType")
    Dim CategoryId As String = Request.QueryString("CategoryId")
    Dim Keyword As String = Web.HttpUtility.UrlDecode(Request.QueryString("Keyword"))
    Dim Sort As String = Request.QueryString("Sort")
    Dim SuccessScript As String = "<script>alert('問題發佈成功');window.location.href='/knowledge/myknowledge_question.aspx?ArticleId=" & ArticleId & _
                                  "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort & "';</script>"

    Dim QFlag As Boolean = False
    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "UPDATE CuDTGeneric SET fCTUPublic = 'Y', xBody = @xbody,Abstract = @Abstract, dEditDate = GETDATE() WHERE iCUItem = @icuitem AND iEditor = @ieditor"

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@xbody", QuestionAdditionalContentText.Text.Trim.Replace("<br />", Chr(13)).Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;"))
      myCommand.Parameters.AddWithValue("@Abstract", "")
	  myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      QFlag = True

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Fail", ScriptFail)
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    REM '---新增問題---加分---(舊的程式碼)
    REM Dim Activity As New ActivityFilter("", MemberId)
    REM Dim ActivityFlag As Boolean = Activity.CheckQuestionGrade

    REM If ActivityFlag Then
      REM Activity.QuestionGradeAdd()
    REM End If

    If QFlag Then

      '---20080927---vincent---kpi share---
      Dim share As New KPIShare(MemberId, "shareAsk", ArticleId)
      share.HandleShare()

      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Success", SuccessScript)

    End If

  End Sub

  Protected Sub QuestionSaveBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QuestionSaveBtn.Click

    Dim ArticleType As String = Request.QueryString("ArticleType")
    Dim CategoryId As String = Request.QueryString("CategoryId")
    Dim Keyword As String = Web.HttpUtility.UrlDecode(Request.QueryString("Keyword"))
    Dim Sort As String = Request.QueryString("Sort")
    Dim SuccessScript As String = "<script>alert('問題暫存成功');window.location.href='/knowledge/myknowledge_question_draft.aspx?ArticleId=" & ArticleId & _
                                  "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort & "';</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "UPDATE CuDTGeneric SET xBody = @xbody, dEditDate = GETDATE() WHERE iCUItem = @icuitem AND iEditor = @ieditor"

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@xbody", QuestionAdditionalContentText.Text.Trim.Replace("<br />", Chr(13)).Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;"))
      myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Success", SuccessScript)

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Fail", ScriptFail)
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  Protected Sub QuestionDelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles QuestionDelBtn.Click

    Dim ArticleType As String = Request.QueryString("ArticleType")
    Dim CategoryId As String = Request.QueryString("CategoryId")
    Dim Keyword As String = Web.HttpUtility.UrlDecode(Request.QueryString("Keyword"))
    Dim Sort As String = Request.QueryString("Sort")
    Dim SuccessScript As String = "<script>alert('問題刪除成功');window.location.href='/knowledge/myknowledge_question_lp.aspx?ArticleId=" & ArticleId & _
                                  "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort & "';</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    '---檢查是否可以刪除---
    Dim Flag As Boolean = False
    myConnection = New SqlConnection(ConnString)
    Try

      sqlString = "SELECT iCUItem FROM CuDTGeneric WHERE fCTUPublic = 'N' AND iEditor = @ieditor AND iCUItem = @icuitem"
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)
      myReader = myCommand.ExecuteReader
      If myReader.HasRows Then
        Flag = True
      Else
        Flag = False
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Fail", "<script>alert('此篇文章無法被刪除!!');history.back();</script>")
        Exit Sub
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

      If Flag Then

        sqlString = "UPDATE KnowledgeForum SET Status = 'D' WHERE gicuitem = @icuitem "

        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@icuitem", ArticleId)

        myCommand.ExecuteNonQuery()
        myCommand.Dispose()

        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Success", SuccessScript)
      End If
    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "Fail", ScriptFail)
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

End Class
