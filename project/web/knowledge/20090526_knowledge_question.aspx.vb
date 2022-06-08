Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_question
  Inherits System.Web.UI.Page

  Dim iBaseDSD As String = WebConfigurationManager.AppSettings("KnowledgeBaseDSD")
  Dim iCtUnit As String = WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId")
  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sqlString As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim keywords As String = ""
  Dim SuccessScript As String = ""
  Dim MaxIcuitem As String = ""
  Dim subjectmp As String = "" '主題館發問mp
  Dim myReader As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim MemberId As String
        Dim Script As String = ""
        Script = "<script>"
        If Request.QueryString("type") = "logout" Then
            Script &= "window.location.href='/knowledge/knowledge.aspx';"
        Else
            'Script &= "alert('連線逾時或尚未登入，請登入會員');window.location.href='" & Request.UrlReferrer.ToString() &"';"
			Script &= "alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';"
        End If 
        Script &= "</script>"
        '-----------------------------------------------------------------------------
        If Session("memID") = "" Or Session("memID") Is Nothing Then
            MemberId = ""
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
            Exit Sub
        Else
            MemberId = Session("memID").ToString()

            '---插入可上傳圖片的數量---
            CheckMemberUploadPic(MemberId)

            'Try
            '  If Request.QueryString("xItem") <> "" And Request.QueryString("CtNode") <> "" Then
            '    Dim ReUnit As New xUnitPath(Request.QueryString("xItem"), Request.QueryString("CtNode"), Request.QueryString("mp"))
            '    RelationUnit.Text = "<div class=""relcontent"">"
            '    '---first : mp value---
            '    RelationUnit.Text &= ReUnit.getmpPath
            '    '---second : CtNode value---
            '    RelationUnit.Text &= ReUnit.getCtNodePath()
            '    '---third : xItem value---
            '    RelationUnit.Text &= ReUnit.getxItemPath
            '    RelationUnit.Text &= "</div>"
            '  End If
            'Catch ex As Exception
            '  Response.Write(ex.ToString)
            'End Try
			'主題館發問 標題預設值
			subjectmp = Request.QueryString("subjectmp")
			myConnection = New SqlConnection(ConnString)
			if subjectmp <> "" then			
				sqlString = "SELECT xdmpName FROM xdmpList WHERE (xdmpID = @xdmpID) "
				myConnection.Open()
				myCommand = New SqlCommand(sqlString, myConnection)
				myCommand.Parameters.AddWithValue("@xdmpID", subjectmp)
				myReader = myCommand.ExecuteReader()
				if myReader.read() then
					if QuestionTitle.Text = "請輸入清楚明白的問題標題，限 30 個字以內" then
						QuestionTitle.Text = "《" & myReader("xdmpName") & "》"
						Tag1.Text = myReader("xdmpName")
					end if
				end if
				myReader.Close()
				myCommand.Dispose()
			end if
			'主題館發問 標題預設值end
        End If
        '-----------------------------------------------------------------------------
		
    End Sub

  Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SubmitBtn.Click

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
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');history.back();</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    If QuestionTitle.Text = "請輸入清楚明白的問題標題，限 30 個字以內" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", "<script>alert('請輸入問題標題!!');history.back();</script>")
      Exit Sub
    End If
    If QuestionContent.Text = "請詳細說明問題的狀況" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", "<script>alert('請輸入問題內容!!');history.back();</script>")
      Exit Sub
    End If

    Dim QFlag As Boolean = False

    HandleTagContent(keywords)

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, topCat, xKeyword, xBody, siteId) "
      sqlString &= "VALUES(@ibasedsd, @ictunit, 'Y', @stitle, @ieditor, '0', @topcat, @xkeyword, @xbody, @siteid)"

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.Add(New SqlParameter("@ibasedsd", iBaseDSD))
      myCommand.Parameters.Add(New SqlParameter("@ictunit", iCtUnit))
      myCommand.Parameters.Add(New SqlParameter("@stitle", QuestionTitle.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")))
      myCommand.Parameters.Add(New SqlParameter("@ieditor", MemberId))
      myCommand.Parameters.Add(New SqlParameter("@topcat", QuestionTypeDDL.SelectedValue))
      myCommand.Parameters.Add(New SqlParameter("@xkeyword", keywords))
      myCommand.Parameters.Add(New SqlParameter("@xbody", QuestionContent.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")))
      myCommand.Parameters.Add(New SqlParameter("@siteid", WebConfigurationManager.AppSettings("SiteId")))

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
	  

      sqlString = "SELECT @@identity AS MaxIcuitem"
      Dim MaxCommand As SqlDataAdapter = New SqlDataAdapter(sqlString, myConnection)
      Dim MaxDataSet As DataSet = New DataSet()
      MaxCommand.Fill(MaxDataSet, "IndexRecommend")
      Dim myTable As DataTable = MaxDataSet.Tables("IndexRecommend")
      MaxIcuitem = myTable.Rows(0).Item(0)
	  
	   '主題館發問
	if subjectmp <>"" then
	  Dim xurl as string =""
	  Dim baseDSDId as string =""
	  Dim unitId as string =""
	  Dim xurl2 as string =""
	  xurl2 = "/Knowledge/Knowledge_cp.aspx?ArticleId=" & MaxIcuitem & "&ArticleType=A&CategoryId=" & QuestionTypeDDL.SelectedValue
		sqlString = "SELECT baseDSDId,ctUnitId FROM KnowledgeToSubject WHERE status='k' and subjectId = @subjectId "
		myCommand = New SqlCommand(sqlString, myConnection)
		myCommand.Parameters.AddWithValue("@subjectId", subjectmp)
		myReader = myCommand.ExecuteReader()
		if myReader.read() then
			baseDSDId = myReader("baseDSDId")
			unitId = myReader("ctUnitId")
		end if
		myReader.Close()
		myCommand.Dispose()
	  '主表
	  sqlString = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, xURL, showType, siteId, xNewWindow) "
	  sqlString &= " VALUES(@baseDSDId,@unitId, 'Y',@stitle, @ieditor, '0',@xurl, '2', '2', 'Y') "

      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.Add(New SqlParameter("@baseDSDId", baseDSDId))
      myCommand.Parameters.Add(New SqlParameter("@unitId", unitId))
      myCommand.Parameters.Add(New SqlParameter("@stitle", QuestionTitle.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")))
      myCommand.Parameters.Add(New SqlParameter("@ieditor", MemberId))
	  myCommand.Parameters.Add(New SqlParameter("@xurl", xurl2))

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
	  '副表
	  sqlString = "INSERT INTO CuDTx7 (gicuitem) VALUES(@gicuitem)"
	  myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.Add(New SqlParameter("@gicuitem", MaxIcuitem+1))
	  myCommand.ExecuteNonQuery()
      myCommand.Dispose()
	end if
	  '主題館發問 end

      sqlString = "INSERT INTO KnowledgeForum(gicuitem) VALUES (@gicuitem)"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.Add(New SqlParameter("@gicuitem", MaxIcuitem))
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      Dim uploadedFiles As HttpFileCollection = Request.Files
      For i As Integer = 0 To uploadedFiles.Count - 1
        Dim myfileupload = CType(filePanel.FindControl("picFile_" & i), FileUpload)
        Dim myfileuploadname As TextBox = CType(filePanel.FindControl("picDesc_" & i), TextBox)

        If myfileupload.HasFile Then
          Try
            Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)  
            fileExtension = IO.Path.GetExtension(myfileupload.FileName).ToLower   '取得檔案格式 
            Dim filenane As String

            filenane = Now.Year() & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & Now.Millisecond & i
            fileExtension = filenane & fileExtension
            myfileupload.SaveAs(Server.MapPath("/public/Data/") & fileExtension)

            sqlString = "INSERT INTO KnowledgePicInfo (picTitle, picPath, picStatus, parentIcuitem, picType, picOrder)"
            sqlString &= "VALUES(@stitle,@route,'W',@parentIcuitem,'A',@picOrder)"

            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.Add(New SqlParameter("@stitle", myfileuploadname.Text))
            myCommand.Parameters.Add(New SqlParameter("@route", "/public/Data/" & fileExtension))
            myCommand.Parameters.Add(New SqlParameter("@parentIcuitem", MaxIcuitem))
            myCommand.Parameters.Add(New SqlParameter("@picOrder", i + 1))

            myCommand.ExecuteNonQuery()
            myCommand.Dispose()
          Catch ex As Exception
            Response.Write(ex)

          End Try

        End If

      Next

      QFlag = True

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Response.Redirect("/knowledge/knowledge_question.aspx")
      'Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Dim Activity As New ActivityFilter("", MemberId)
    Dim ActivityFlag As Boolean = Activity.CheckQuestionGrade

    If ActivityFlag Then
      Activity.QuestionGradeAdd()
    End If

    If QFlag Then

      '---20080927---vincent---kpi share---
      Dim share As New KPIShare(MemberId, "shareAsk", MaxIcuitem)
      share.HandleShare()

      SuccessScript = "<script>alert('發問成功');window.location.href='/knowledge/myknowledge_question_lp.aspx';</script>"
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "submit", SuccessScript)

    End If

  End Sub

  Protected Sub TempSubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TempSubmitBtn.Click

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
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');history.back();</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------
   
    Dim SuccessScript As String = ""
    Dim MaxIcuitem As String = ""

    HandleTagContent(keywords)

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, topCat, xKeyword, xBody, siteId) "
      sqlString &= "VALUES(@ibasedsd, @ictunit, 'N', @stitle, @ieditor, '0', @topcat, @xkeyword, @xbody, @siteid)"

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.Add(New SqlParameter("@ibasedsd", iBaseDSD))
      myCommand.Parameters.Add(New SqlParameter("@ictunit", iCtUnit))
      myCommand.Parameters.Add(New SqlParameter("@stitle", QuestionTitle.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")))
      myCommand.Parameters.Add(New SqlParameter("@ieditor", MemberId))
      myCommand.Parameters.Add(New SqlParameter("@topcat", QuestionTypeDDL.SelectedValue))
      myCommand.Parameters.Add(New SqlParameter("@xkeyword", keywords))
      myCommand.Parameters.Add(New SqlParameter("@xbody", QuestionContent.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")))
      myCommand.Parameters.Add(New SqlParameter("@siteid", WebConfigurationManager.AppSettings("SiteId")))

      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sqlString = "SELECT @@identity AS MaxIcuitem"
      Dim MaxCommand As SqlDataAdapter = New SqlDataAdapter(sqlString, myConnection)
      Dim MaxDataSet As DataSet = New DataSet()
      MaxCommand.Fill(MaxDataSet, "IndexRecommend")
      Dim myTable As DataTable = MaxDataSet.Tables("IndexRecommend")
      MaxIcuitem = myTable.Rows(0).Item(0)

      sqlString = "INSERT INTO KnowledgeForum(gicuitem) VALUES (@gicuitem)"
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.Add(New SqlParameter("@gicuitem", MaxIcuitem))
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      SuccessScript = "<script>alert('發問暫存成功');window.location.href='/knowledge/myknowledge_question_lp.aspx';</script>"
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "submit", SuccessScript)

    Catch ex As Exception
      If Request.QueryString("debug") = "true" Then
        Response.Write(ex.ToString)
        Response.End()
      End If
      Response.Redirect("/knowledge/knowledge_question.aspx")
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  Protected Sub ResetBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ResetBtn.Click

    Response.Redirect(Request.QueryString("BackUrl"))

  End Sub

  Sub CheckMemberUploadPic(ByVal memberId As String)

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    'Dim myReader As SqlDataReader
    Dim uploadRight As String = ""
    Dim picCount As Integer = 0

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "SELECT ISNULL(uploadRight, '') AS uploadRight, ISNULL(uploadPicCount, 0) AS UploadCount FROM Member WHERE account = @memberid"
      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        uploadRight = myReader("uploadRight")
        picCount = myReader("UploadCount")
      End If
      If Not myReader.IsClosed Then myReader.Close()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try

    If uploadRight <> "Y" Then
      uploadPicDiv.Visible = False
    Else
      If picCount = 0 Then
        uploadPicDiv.Visible = False
      Else
        Dim picFile As FileUpload
        Dim picDesc As TextBox
        Dim picLabel As Label

        uploadPicDiv.Visible = True
        For i As Integer = 0 To picCount - 1

          picLabel = New Label
          picLabel.Text = "<li>"
          filePanel.Controls.Add(picLabel)

          picFile = New FileUpload
          picFile.ID = "picFile_" & i
          picFile.Attributes.Add("onchange", "checkFileSize(" & i & ")")
          filePanel.Controls.Add(picFile)

          picLabel = New Label
          picLabel.Text = "<br />"
          filePanel.Controls.Add(picLabel)

          picDesc = New TextBox
          picDesc.ID = "picDesc_" & i
          picDesc.Text = "請輸入圖片說明"
          picDesc.Attributes.Add("onclick", "this.value='';")
          filePanel.Controls.Add(picDesc)

          picLabel = New Label
          picLabel.Text = "</li>"
          filePanel.Controls.Add(picLabel)

        Next
      End If
    End If
    If uploadPicDiv.Visible = True Then
      SubmitBtn.OnClientClick = "return checkFile('" & picCount & "')"
    End If

  End Sub

  Sub HandleTagContent(ByRef keywords)

    If Tag1.Text.Trim <> "" Then
      keywords &= Tag1.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If
    If Tag2.Text.Trim <> "" Then
      keywords &= Tag2.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If
    If Tag3.Text.Trim <> "" Then
      keywords &= Tag3.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If
    If Tag4.Text.Trim <> "" Then
      keywords &= Tag4.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If
    If Tag5.Text.Trim <> "" Then
      keywords &= Tag5.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If
    If Tag6.Text.Trim <> "" Then
      keywords &= Tag6.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If
    If Tag7.Text.Trim <> "" Then
      keywords &= Tag7.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If
    If Tag8.Text.Trim <> "" Then
      keywords &= Tag8.Text.Trim.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;") & ";"
    End If

  End Sub

End Class
