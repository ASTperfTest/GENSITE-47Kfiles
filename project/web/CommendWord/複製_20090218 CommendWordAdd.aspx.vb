Imports System
Imports System.Data
Imports System.Data.SqlClient

Partial Class CommendWord_CommendWordAdd
  Inherits System.Web.UI.Page
  Implements System.Web.UI.ICallbackEventHandler

  Dim CallbackResult As String = "Y"
  Dim bOK As Boolean

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim MemberId As String
    Dim Script As String = "<script>alert('若要推薦詞彙,請先登入會員!!');window.close();</script>"
    Dim ScriptWord As String = "<script>alert('推薦詞彙參數錯誤!!');window.close();</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    Dim Word As String = Request.QueryString("word")
    If Word <> "" Then
      'Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoWord", ScriptWord)
      'Exit Sub
      'Else
      WordTitle.Text = Word
    End If

    If Request.QueryString("type") = "0" Then
      '---入口網---
      PathText.Text = GetPath("0", Request.QueryString("xItem"), Request.QueryString("ctNode"), Request.QueryString("mp"))

    ElseIf Request.QueryString("type") = "1" Then
      '---知識庫---
      PathText.Text = ListCategory(Request.QueryString("ReportId"), Request.QueryString("DatabaseId"), Request.QueryString("CategoryId"), Request.QueryString("ActorType"))

    ElseIf Request.QueryString("type") = "2" Then
      '---主題館---
      PathText.Text = GetPath("2", Request.QueryString("xItem"), Request.QueryString("ctNode"), Request.QueryString("mp"))

    End If

    errorshow.Visible = "false"
    bOK = False
    If (Session("CaptchaImageText") = Nothing) Then
    ElseIf (CodeNumberTextBox.Text = Session("CaptchaImageText").ToString()) Then
      bOK = True
    Else
    End If

    Dim cbReference As String = Page.ClientScript.GetCallbackEventReference(Me, "arg", "ReceiveServerData", "context")
    Dim cbScript As String = "function CallServer(arg, context) {" & cbReference & "}"
    Page.ClientScript.RegisterClientScriptBlock(Me.GetType(), "CallServer", cbScript, True)

  End Sub

  Public Function GetCallbackResult() As String Implements System.Web.UI.ICallbackEventHandler.GetCallbackResult

    Return CallbackResult

  End Function

  Public Sub RaiseCallbackEvent(ByVal word As String) Implements System.Web.UI.ICallbackEventHandler.RaiseCallbackEvent

    If Not CheckIsWordDuplicate(word) Then
      CallbackResult = "Y"
    Else
      CallbackResult = "N"
    End If

  End Sub

  Function CheckIsWordDuplicate(ByVal word As String) As Boolean

    Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim count As Integer = 0

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "SELECT COUNT(*) FROM CuDTGeneric WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.sTitle = @word)"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", ConfigurationManager.AppSettings("PediaCtUnitId"))
      myCommand.Parameters.AddWithValue("@word", word)
      count = myCommand.ExecuteScalar
      If count = 0 Then
        Return True
      Else
        Return False
      End If
      myCommand.Dispose()

    Catch ex As Exception      
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Function

  Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SubmitBtn.Click

    Dim MemberId As String
    Dim Script As String = "<script>alert('若要推薦詞彙,請先登入會員!!');window.close();</script>"

    Dim path As String = ""
    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If    
    '-----------------------------------------------------------------------------

    If WordTitle.Text.Trim = "" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "dup", "<script>alert('推薦的詞彙不可空白!!');window.close();</script>")
      Exit Sub
    End If

    If Not CheckIsWordDuplicate(WordTitle.Text) Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "dup", "<script>alert('推薦的詞彙重複!!');window.close();</script>")
      Exit Sub
    End If

    If Request.QueryString("type") = "0" Then
      '---入口網---
      path = "/ct.asp?xItem=" & Request.QueryString("xItem") & "&ctNode=" & Request.QueryString("ctNode") & "&mp=" & Request.QueryString("mp")
    ElseIf Request.QueryString("type") = "1" Then
      '---知識庫---
      path = "/CatTree/CatTreeContent.aspx?ReportId=" & Request.QueryString("ReportId") & "&DatabaseId=" & Request.QueryString("DatabaseId") & "&CategoryId=" & Request.QueryString("CategoryId") & "&ActorType=" & Request.QueryString("ActorType")
    ElseIf Request.QueryString("type") = "2" Then
      '---主題館---
      path = "/subject/ct.asp?xItem=" & Request.QueryString("xItem") & "&ctNode=" & Request.QueryString("ctNode") & "&mp=" & Request.QueryString("mp")
    End If

    Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand

    Dim maxIcuitem As String = ""

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, topCat, vGroup, xKeyword, xBody, siteId) "
      sql &= "VALUES(@ibasedsd, @ictunit, 'N', @stitle, @ieditor, '0', '', 'Y', '', '', @siteid)"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@ibasedsd", ConfigurationManager.AppSettings("PediaBaseDSD"))
      myCommand.Parameters.AddWithValue("@ictunit", ConfigurationManager.AppSettings("PediaCtUnitId"))
      myCommand.Parameters.AddWithValue("@stitle", WordTitle.Text)
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)
      myCommand.Parameters.AddWithValue("@siteid", ConfigurationManager.AppSettings("PediaSiteId"))

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sql = "SELECT @@identity AS MaxIcuitem"
      Dim MaxCommand As SqlDataAdapter = New SqlDataAdapter(sql, myConnection)
      Dim MaxDataSet As DataSet = New DataSet()
      MaxCommand.Fill(MaxDataSet, "commendIndex")
      Dim myTable As DataTable = MaxDataSet.Tables("commendIndex")
      maxIcuitem = myTable.Rows(0).Item(0)

      sql = "INSERT INTO Pedia(gicuitem, engTitle, formalName, localName, memberId, commendTime, quoteIcuitem, parentIcuitem, xStatus, path) "
      sql &= " VALUES (@gicuitem, '', '', '', @memberId, GETDATE(), 0, 0, 'Y', @path)"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@gicuitem", maxIcuitem)
      myCommand.Parameters.AddWithValue("@memberId", MemberId)
      myCommand.Parameters.AddWithValue("@path", path)

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

    Catch ex As Exception
      'Response.Write(ex.ToString)
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    CheckPediaActivity(MemberId)

    '---20080927---vincent---kpi share---
    Dim share As New KPIShare(MemberId, "shareSuggest", maxIcuitem)
    share.HandleShare()
    '------------------------------------

    Dim ScriptWord As String = "<script>alert('詞彙推薦成功，本詞彙需待管理者進行審核!!');window.close();</script>"

    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "success", ScriptWord)


  End Sub

  Sub CheckPediaActivity(ByVal memberid As String)

    Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    myConnection = New SqlConnection(ConnString)
    Try
      Dim pediaOpen As Boolean = False
      Dim pediaFlag As Boolean = False

      sql = "SELECT * FROM Activity WHERE ActivityId = 'pedia' AND GETDATE() BETWEEN ActivityStartTime AND ActivityEndTime"
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        pediaOpen = True
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

      '---活動開放期間---
      If pediaOpen Then

        sql = "SELECT * FROM ActivityPediaMember WHERE memberId = @memberid"
        myCommand = New SqlCommand(sql, myConnection)
        myCommand.Parameters.AddWithValue("@memberid", memberid)
        myReader = myCommand.ExecuteReader
        If myReader.Read Then
          pediaFlag = True
        End If
        If Not myReader.IsClosed Then
          myReader.Close()
        End If
        myCommand.Dispose()

        If pediaFlag Then
          'sql = "UPDATE ActivityPediaMember SET commendAdditionalCount = commendAdditionalCount + 1 WHERE memberId = @memberid"
          'myCommand = New SqlCommand(sql, myConnection)
          'myCommand.Parameters.AddWithValue("@memberid", memberid)
          'myCommand.ExecuteNonQuery()
          'myCommand.Dispose()
        Else
          sql = "INSERT INTO ActivityPediaMember VALUES(@memberid, 0, 0)"
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@memberid", memberid)
          myCommand.ExecuteNonQuery()
          myCommand.Dispose()
        End If

      End If

    Catch ex As Exception
      Response.Write(ex.ToString)
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  Public Function ListCategory(ByVal reportid As String, ByVal databaseid As String, ByVal categoryid As String, ByVal actortype As String) As String

    Dim ConnString As String = ConfigurationManager.ConnectionStrings("KMConnString").ConnectionString
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim sqlstring As String
    Dim sb As New StringBuilder
    Dim myReader1 As SqlDataReader

    myConnection = New SqlConnection(ConnString)
    Try
      sqlstring = "SELECT DISTINCT CATEGORY.CATEGORY_ID AS CatId1, CATEGORY.CATEGORY_NAME AS CatName1, "
      sqlstring &= "CATEGORY_1.CATEGORY_ID as CatId2, CATEGORY_1.CATEGORY_NAME AS CatName2, "
      sqlstring &= "CATEGORY_2.CATEGORY_ID as CatId3, CATEGORY_2.CATEGORY_NAME AS CatName3 "
      sqlstring &= "FROM CATEGORY AS CATEGORY_2 RIGHT OUTER JOIN CATEGORY AS CATEGORY_1 ON CATEGORY_2.CATEGORY_ID = CATEGORY_1.PARENT_CATEGORY_ID "
      sqlstring &= "RIGHT OUTER JOIN REPORT INNER JOIN CAT2RPT ON REPORT.REPORT_ID = CAT2RPT.REPORT_ID INNER JOIN "
      sqlstring &= "CATEGORY ON CAT2RPT.CATEGORY_ID = CATEGORY.CATEGORY_ID ON CATEGORY_1.CATEGORY_ID = CATEGORY.PARENT_CATEGORY_ID "
      sqlstring &= "WHERE (CATEGORY.DATA_BASE_ID = 'DB020') AND (REPORT.REPORT_ID = '" & reportid & "') ORDER BY CatId1"

      myConnection.Open()
      myCommand = New SqlCommand(sqlstring, myConnection)
      myReader1 = myCommand.ExecuteReader()

      sb.AppendLine("首頁 ")

      While myReader1.Read()

        If myReader1.Item("CatId3") IsNot DBNull.Value Then
          sb.AppendLine("&gt; " & myReader1("CatName3"))
        End If
        If myReader1.Item("CatId2") IsNot DBNull.Value Then
          sb.AppendLine("&gt; " & myReader1("CatName2"))
        End If
        If myReader1.Item("CatId1") IsNot DBNull.Value Then
          sb.AppendLine("&gt; " & myReader1("CatName1"))
        End If

      End While

      If Not myReader1.IsClosed Then
        myReader1.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      'Response.Write(ex.ToString)
      Response.Redirect("/mp.asp?mp=1")
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return sb.ToString

  End Function


  Function GetPath(ByVal type As String, ByVal xitem As String, ByVal ctnode As String, ByVal mp As String) As String

    Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim sqlstring As String
    Dim sb As New StringBuilder
    Dim myReader1 As SqlDataReader
    myConnection = New SqlConnection(ConnString)
    Try
      sqlstring = " Select CatName FROM CatTreeNode WHERE (CtNodeID = @nodeid)"

      myConnection.Open()
      myCommand = New SqlCommand(sqlstring, myConnection)
      myCommand.Parameters.AddWithValue("@nodeid", ctnode)
      myReader1 = myCommand.ExecuteReader()

      sb.AppendLine("首頁 ")

      While myReader1.Read()

        sb.AppendLine("&gt; " & myReader1("CatName"))

      End While

      If Not myReader1.IsClosed Then
        myReader1.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      'Response.Write(ex.ToString)
      Response.Redirect("/mp.asp?mp=1")
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return sb.ToString

  End Function

End Class
