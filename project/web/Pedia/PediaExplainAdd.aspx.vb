Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Pedia_PediaExplainAdd
    Inherits System.Web.UI.Page

  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sql As String
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader
  Dim Title As String
    Dim maxIcuitem As String = ""
    Protected guid As String = ""

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim aid As String = "0"
    Dim paid As String = "0"
    If Request.QueryString("AId") <> "" Then
            aid = Request.QueryString("AId")
            'sam 弱掃
            If (WebUtility.checkInt(aid) = False) Then
                Response.Redirect("/mp.asp")
            End If
        End If
        If Request.QueryString("PAId") <> "" Then
            paid = Request.QueryString("PAId")
            If (WebUtility.checkInt(paid) = False) Then
                Response.Redirect("/mp.asp")
            End If
        End If

        Dim MemberId As String = ""
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='" & Request.UrlReferrer.ToString() & "';</script>"
        guid = System.Guid.NewGuid.ToString
        guid = guid.Replace("-", "")
        MemberCaptChaImage.ImageUrl = "~/CaptchaImage/JpegImage.aspx?guid=" & guid
        '-----------------------------------------------------------------------------
        If Session("memID") = "" Or Session("memID") Is Nothing Then
            MemberId = ""
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
            Exit Sub
        Else
            MemberId = Session("memID").ToString()
        End If

        If Not FillPediaContent(paid) Then
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", "<script>alert('此詞目不提供補充!!');history.back();</script>")
            Exit Sub
        End If

        If Not Page.IsPostBack Then

            If aid = "0" Then
                '---詞目內容頁補充---
                If Not CheckHaveExplain(paid) Then
                    '---若沒有補充,第一筆, 內容預設空白---
                    ExplainContent.Text = ""
                Else
                    '---若有補充,引用最新的解釋---
                    ExplainContent.Text = GetLatestExplain(paid)
                End If
            Else
                '---解釋內容頁補充---
                '---預設帶這一筆內容---
                ExplainContent.Text = GetCurrentExplain(aid)
            End If

        End If

  End Sub

  Function FillPediaContent(ByVal paid As String) As Boolean

    Dim flag As Boolean = True
    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT CuDTGeneric.sTitle, Pedia.engTitle, CuDTGeneric.vGroup FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (CuDTGeneric.iCUItem = @icuitem)"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", paid)
      myReader = myCommand.ExecuteReader

      If myReader.Read Then
        If myReader("vGroup") IsNot DBNull.Value Then
          If myReader("vGroup").ToString <> "Y" Then
            flag = False
          End If
        End If
        Title = myReader("sTitle")
        WordTitle.Text = "<a href=""javascript:WordSearch('" & myReader("sTitle") & "')"" target=""_blank"">" & myReader("sTitle") & "</a>"
        WordEngTitle.Text = IIf(myReader("engTitle") Is DBNull.Value Or myReader("engTitle") = "", "&nbsp;", myReader("engTitle"))
      Else
        Session("PediaErrMsg") = "詞目不存在"
        ShowErrorMessage()
        flag = False
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

      Return flag

    Catch ex As Exception
      Session("PediaErrMsg") = "詞目錯誤"
      ShowErrorMessage()
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Function

  Function CheckHaveExplain(ByVal paid As String) As Boolean

    Dim count As Integer = 0

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT COUNT(*) FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (Pedia.parentIcuItem = @icuitem) "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", paid)
      count = myCommand.ExecuteScalar

      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "檢查解釋錯誤"
      Exit Function
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    If count = 0 Then
      Return False
    Else
      Return True
    End If

  End Function

  Function GetCurrentExplain(ByVal aid As String) As String

    Dim str As String = ""
    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT xBody FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (CuDTGeneric.iCUItem = @icuitem) "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", aid)
      myReader = myCommand.ExecuteReader
      If myReader.Read() Then
        str = myReader("xBody")
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "檢查解釋錯誤"
      Return ""
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return str

  End Function

  Function GetLatestExplain(ByVal paid As String) As String

    Dim str As String = ""
    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT TOP(1) xBody FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (Pedia.parentIcuitem = @icuitem) "
      sql &= " ORDER BY Pedia.commendTime DESC "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", paid)
      myReader = myCommand.ExecuteReader
      If myReader.Read() Then
        str = myReader("xBody")
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "檢查解釋錯誤"
      Return ""
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return str

  End Function

  Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SubmitBtn.Click
        Dim oldGuid = ""
        If Not (Request.Form("MemberGuid").ToString() = "") Then
            oldGuid = Request.Form("MemberGuid").ToString()
        End If
    Dim PicInput As Boolean = False
    '---handle picture---
        If (Session(oldGuid & "CaptchaImageText") = Nothing) Then
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼失效!!');window.location.href='" & Request.RawUrl() & "';</script>")
            Exit Sub
        ElseIf (CaptChaText.Text <> Session(oldGuid & "CaptchaImageText").ToString()) Then
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼錯誤!!');history.back();</script>")
            Exit Sub
        End If
        Session.Remove(oldGuid & "CaptchaImageText")
    Dim MemberId As String = Session("memID").ToString()

    Dim aid As String = "0"
    Dim paid As String = "0"
    If Request.QueryString("AId") <> "" Then
      aid = Request.QueryString("AId")
    End If
    If Request.QueryString("PAId") <> "" Then
      paid = Request.QueryString("PAId")
    End If
    '---檢查內容是否有修改---
    If Not CheckIsBodyModify(aid, paid) Then
      ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('請修改補充解釋內容!!');history.back();</script>")
    Else

      If Not InsertBodyIntoDB(aid, paid, MemberId, "Y") Then
        ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('補充解釋錯誤!!');history.back();</script>")
      Else
        CheckPediaActivity(MemberId)
        '---20080927---vincent---kpi share---
        Dim share As New KPIShare(MemberId, "shareExplain", maxIcuitem)
        share.HandleShare()
        '------------------------------------
        ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('補充解釋成功!!');window.location.href='/Pedia/PediaContent.aspx?AId=" & paid & "';</script>")
      End If

    End If

  End Sub

  Protected Sub TempSaveBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TempSaveBtn.Click
        Dim oldGuid = ""
        If Not (Request.Form("MemberGuid").ToString() = "") Then
            oldGuid = Request.Form("MemberGuid").ToString()
        End If
    Dim PicInput As Boolean = False
    '---handle picture---
        If (Session(oldGuid & "CaptchaImageText") = Nothing) Then
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼失效!!');window.location.href='" & Request.RawUrl() & "';</script>")
            Exit Sub
        ElseIf (CaptChaText.Text <> Session(oldGuid & "CaptchaImageText").ToString()) Then
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼錯誤!!');history.back();</script>")
            Exit Sub
        End If
        Session.Remove(oldGuid & "CaptchaImageText")

    Dim MemberId As String = Session("memID").ToString()

    Dim aid As String = "0"
    Dim paid As String = "0"
    If Request.QueryString("AId") <> "" Then
      aid = Request.QueryString("AId")
    End If
    If Request.QueryString("PAId") <> "" Then
      paid = Request.QueryString("PAId")
    End If
    '---檢查內容是否有修改---
    If Not CheckIsBodyModify(aid, paid) Then
      ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('請修改補充解釋內容!!');history.back();</script>")
    Else

      If Not InsertBodyIntoDB(aid, paid, MemberId, "N") Then
        ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('補充解釋錯誤!!');history.back();</script>")
      Else
        CheckPediaActivity(MemberId)
        ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('補充解釋成功!!');window.location.href='/Pedia/PediaContent.aspx?AId=" & paid & "';</script>")
      End If

    End If

  End Sub

  Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click

    Response.Redirect("/Pedia/PediaExplainAdd.aspx?AId=" & Request.QueryString("AId") & "&PAId=" & Request.QueryString("PAId"))

  End Sub

  Function CheckIsBodyModify(ByVal aid As String, ByVal paid As String) As Boolean

    If ExplainContent.Text = "" Then
      Return False
    End If

    If aid = "0" Then
      If Not CheckHaveExplain(paid) Then
        '---若沒有補充,第一筆, 內容預設空白---
        Return True
      Else
        '---若有補充,引用最新的解釋---
        If ExplainContent.Text = GetLatestExplain(paid) Then
          Return False
        Else
          Return True
        End If
      End If
    Else
      '---預設帶這一筆內容---
      If ExplainContent.Text = GetCurrentExplain(aid) Then
        Return False
      Else
        Return True
      End If
    End If
  End Function

  Function InsertBodyIntoDB(ByVal aid As String, ByVal paid As String, ByVal memberId As String, ByVal xStatus As String) As Boolean

    Dim explain As String = ExplainContent.Text.Trim
    explain = explain.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, xBody, siteId) "
      sql &= "VALUES(@ibasedsd, @ictunit, 'Y', @stitle, @ieditor, '0', @xbody, @siteid)"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@ibasedsd", ConfigurationManager.AppSettings("PediaBaseDSD"))
      myCommand.Parameters.AddWithValue("@ictunit", ConfigurationManager.AppSettings("PediaAdditionalCtUnitId"))
      If explain.Length > 15 Then
        myCommand.Parameters.AddWithValue("@stitle", "補充解釋-" & Title & "-" & paid & "-" & explain.Substring(0, 15))
      Else
        myCommand.Parameters.AddWithValue("@stitle", "補充解釋-" & Title & "-" & paid & "-" & explain)
      End If
      myCommand.Parameters.AddWithValue("@ieditor", memberId)
      myCommand.Parameters.AddWithValue("@xbody", explain)
      myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("PediaSiteId"))

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sql = "SELECT @@identity AS MaxIcuitem"
      Dim MaxCommand As SqlDataAdapter = New SqlDataAdapter(sql, myConnection)
      Dim MaxDataSet As DataSet = New DataSet()
      MaxCommand.Fill(MaxDataSet, "commendIndex")
      Dim myTable As DataTable = MaxDataSet.Tables("commendIndex")
      maxIcuitem = myTable.Rows(0).Item(0)

      sql = "INSERT INTO Pedia(gicuitem, memberId, commendTime, quoteIcuitem, parentIcuitem, xStatus) "
      sql &= " VALUES (@gicuitem, @memberId, GETDATE(), @quoteicuitem, @parenticuitem, @xstatus)"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@gicuitem", maxIcuitem)
      myCommand.Parameters.AddWithValue("@memberId", memberId)
      myCommand.Parameters.AddWithValue("@quoteicuitem", aid)
      myCommand.Parameters.AddWithValue("@parenticuitem", paid)
      myCommand.Parameters.AddWithValue("@xstatus", xStatus)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

    Catch ex As Exception
      'Response.Write(ex.ToString)
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Return False
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return True

  End Function

  Sub CheckPediaActivity(ByVal memberid As String)

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

        '---table中已有這個人, 使用更新---
        If pediaFlag Then
          sql = "UPDATE ActivityPediaMember SET commendAdditionalCount = commendAdditionalCount + 1 WHERE memberId = @memberid"
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@memberid", memberid)
          myCommand.ExecuteNonQuery()
          myCommand.Dispose()
        Else
          sql = "INSERT INTO ActivityPediaMember VALUES(@memberid, 0, 1)"
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@memberid", memberid)
          myCommand.ExecuteNonQuery()
          myCommand.Dispose()
        End If

      End If

    Catch ex As Exception
      'Response.Write(ex.ToString)
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try
    
  End Sub

  Sub ShowErrorMessage()

    Dim str = "<script>alert('" & Session("PediaErrMsg") & "');window.location.href='/Pedia/PediaList.aspx';</script>"

    Session("PediaErrMsg") = ""

    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "PediaError", str)

  End Sub

End Class
