Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_discuss
    Inherits System.Web.UI.Page
    Protected guid As String = ""
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim MemberId As String
        Dim Script As String = ""
        Script = "<script>"
        If Request.QueryString("type") = "logout" Then
            Script &= "window.location.href='/knowledge/knowledge.aspx';"
        Else
            Script &= "alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';"
			'Script &= "alert('連線逾時或尚未登入，請登入會員');window.location.href='" & Request.UrlReferrer.ToString() &"';"
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

        End If
        '-----------------------------------------------------------------------------

        Dim ArticleId As String
        Dim ArticleType As String
        Dim CategoryId As String
        guid = System.Guid.NewGuid.ToString
        guid = guid.Replace("-", "")
        MemberCaptChaImage.ImageUrl = "~/CaptchaImage/JpegImage.aspx?guid=" & guid
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
            CategoryId = "A"
        Else
            CategoryId = Request.QueryString("CategoryId")
        End If

        If Request.QueryString("ArticleId") = "" Then
            Response.Redirect("/knowledge/knowledge.aspx")
            Response.End()
        Else
            ArticleId = Request.QueryString("ArticleId")
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
        BackText.Text = "<a href=""/knowledge/knowledge_discuss_lp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & """ class=""Back"">回上一頁</a>"

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim sb As New StringBuilder
        Dim tagarray As Array
		
		Dim Activity As New ActivityFilter(WebConfigurationManager.AppSettings("ActivityId"), "")
		Dim ActivityFlag As Boolean = Activity.CheckActivity
	
        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT KnowledgeForum.DiscussCount, KnowledgeForum.CommandCount, KnowledgeForum.BrowseCount, KnowledgeForum.TraceCount, "
            sqlString &= "KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, KnowledgeForum.ParentIcuitem, CuDTGeneric.sTitle, CuDTGeneric.xKeyword, ISNULL(CuDTGeneric.vGroup, '') AS vGroup, "
            sqlString &= "CuDTGeneric.xPostDate, CuDTGeneric.xBody, CuDTGeneric.iCUItem, Member.realname, ISNULL(Member.nickname, '') AS nickname "
            sqlString &= "FROM KnowledgeForum INNER JOIN CuDTGeneric ON "
            sqlString &= "KnowledgeForum.gicuitem = CuDTGeneric.iCUItem INNER JOIN Member ON CuDTGeneric.iEditor = Member.account WHERE (CuDTGeneric.iCUItem = @ArticleId) "
            sqlString &= "AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N') "
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@ArticleId", ArticleId)
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
            myReader = myCommand.ExecuteReader()

            If myReader.Read() Then

                If myReader("xPostDate") IsNot DBNull.Value Then
                    xPostDateText.Text = Date.Parse(myReader("xPostDate")).ToShortDateString()
                End If

                sb.Append("<h3>標題：" & myReader("sTitle") & "</h3>")
                sb.Append("<table class=""subtable"" width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""排版用表格"">")
                sb.Append("<tr><td class=""subleft"" style='width:350px'>")
                'Modify by Max 2011/06/07
                sb.Append("<div class=""problem"">" & ReplaceAndFindKeyword(myReader("xBody") & "&nbsp;</div>"))
                'End
                sb.Append(GenerateQuestionAdditional(myReader("iCUItem")))

                sb.Append(GetPicInfo(myReader("iCUItem"), 300))
                sb.Append("</td>")
                sb.Append("<td class=""subright""><div><ul>")
				
				If ActivityFlag Then
				  If myReader("vGroup") = "A" Then
					sb.Append("<li>本問題為「知識問答你˙我˙他」活動問題")
					sb.Append("<div class=""staricon"">")
								sb.Append("<img class=rating src=""../xslGip/style3/images/good.gif"" align=top>")
					sb.Append("</div></li><br/><br/>")
				  End If
				End If
		
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

                sb.Append("<li>瀏覽數：" & myReader("BrowseCount") & "</li><li>討論數：" & myReader("DiscussCount") & "</li>")
                sb.Append("<li>追蹤數：" & myReader("TraceCount") & "</li><li>意見數：" & myReader("CommandCount") & "</li>")
                sb.Append(ConcatStarString(Integer.Parse(myReader("GradeCount")), Integer.Parse(myReader("GradePersonCount"))))
                sb.Append("</ul></td></tr></table>")

                QuestionText.Text = sb.ToString()

                sb = Nothing

            Else
                Script = "<script>alert('{0}!!');window.location.href='/knowledge/knowledge_cp.aspx?ArticleId=" & ArticleId & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "linkerror", Script.Replace("{0}", "文章連結錯誤!!"))
                Exit Sub
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()
            'If hidFileName.Value <> "" And hidFileName.Value <> "undefined" Then
            '    previewImg.Visible = True
            '    previewImg.ImageUrl = hidFileName.Value
            '    lblImgHint.Text = hidFileContent.Value
            '    uploadBtn.Visible = False
            '    deleteBtn.Visible = True
            'Else
            '    lblImgHint.Text = ""
            '    previewImg.Visible = False
            '    uploadBtn.Visible = True
            '    deleteBtn.Visible = False
            'End If
            ' 2010/8/2 vivian modify - 可上傳多張多圖，依"uploadPicCount"決定
            Dim uploadRight As String = ""
            Dim picCount As Integer = 0

            myConnection = New SqlConnection(ConnString)
            Try
                sqlString = "SELECT ISNULL(uploadRight, '') AS uploadRight, ISNULL(uploadPicCount, 0) AS UploadCount FROM Member WHERE account = @memberid"
                myConnection.Open()
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@memberid", MemberId)
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
            'Response.Write(hidFileName.Value)
            If uploadRight = "Y" And picCount <> 0 Then
                If hidFileName.Value <> "" And hidFileName.Value <> "undefined" Then
                    Dim i As Integer
                    Dim files(), hints() As String
                    files = hidFileName.Value.Split("^")
                    hints = hidFileContent.Value.Split("^")
                    If picCount > files.Length - 1 Then
                        uploadBtn.Visible = True
                        textfield.Visible = True
                    Else
                        uploadBtn.Visible = False
                        textfield.Visible = True
                    End If
                    For Each fileName As String In files
                        If fileName <> "" Then
                            Dim img As New Image
                            Dim lbl As New Label
                            lbl.Text = "<br />" & hints(i) & "<p>"
                            i = i + 1
                            img.ID = "previewImg" + i.ToString()
                            img.ImageUrl = fileName
                            img.Attributes("style") = "max-width:450px"
                            divPreviewImg.Controls.Add(img)
                            divPreviewImg.Controls.Add(lbl)
                        End If
                    Next
                Else
                    lblImgHint.Text = ""
                    previewImg.Visible = False
                    uploadBtn.Visible = True
                    textfield.Visible = True
                End If
            Else
                lblImgHint.Text = ""
                previewImg.Visible = False
                uploadBtn.Visible = False
                textfield.Visible = False
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
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    Return str

  End Function

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
    str &= "<div class=""float"">(" & PersonCount & "人評價)</div></li>"

    Return str

  End Function

  Sub CheckMemberUploadPic(ByVal memberId As String)

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
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
        'If uploadPicDiv.Visible = True Then
        '  SubmitBtn.OnClientClick = "return checkFile('" & picCount & "')"
        'End If

  End Sub

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
        ElseIf (MemberCaptChaTBox.Text <> Session(oldGuid & "CaptchaImageText").ToString()) Then
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼錯誤!!');history.back();</script>")
            Exit Sub
        End If
        Session.Remove(oldGuid & "CaptchaImageText")
    Dim MemberId As String = ""
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='" & Request.UrlReferrer.ToString() &"';</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------
	dim discussContentText = DiscussContent.Text.Trim()
    If discussContentText = "請在這裡輸入您對本問題的討論或想法" or discussContentText = "" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", "<script>alert('請輸入討論內容!!');history.back();</script>")
      Exit Sub
    End If

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
      CategoryId = "A"
    Else
      CategoryId = Request.QueryString("CategoryId")
    End If

    Dim iBaseDSD As String = WebConfigurationManager.AppSettings("KnowledgeBaseDSD")
    Dim iCtUnit As String = WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId")
    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim SuccessScript As String = ""
    Dim MaxIcuitem As String = ""
    Dim KFlag As Boolean = False
    Dim ArticleContent As String = DiscussContent.Text.Trim
    ArticleContent = ArticleContent.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, xBody, siteId) "
      sqlString &= "VALUES(@ibasedsd, @ictunit, 'Y', @stitle, @ieditor, '0', @xbody, @siteid)"

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ibasedsd", iBaseDSD)
      myCommand.Parameters.AddWithValue("@ictunit", iCtUnit)
      myCommand.Parameters.AddWithValue("@stitle", "討論-" & Request.QueryString("ArticleId"))
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
      myCommand.Parameters.AddWithValue("@parenticuitem", Request.QueryString("ArticleId"))
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

            ' 2010/5/26 vivian modify - COAKM10099043上傳圖片裁切功能(FileName存在HiddenField)
            ' 2010/8/2 vivian modify - 可上傳多張多圖，依"uploadPicCount"決定
            Dim i As Integer = 0
            Dim files(), hints() As String
            files = hidFileName.Value.Split("^")
            hints = hidFileContent.Value.Split("^")

            'Dim myfileuploadcontent As String = hidFileContent.value
            For Each fileName As String In files
                If fileName <> "" Then
                    'If hidFileName.value <> "" Then
                    Dim sourcePath As String = Server.MapPath("/knowledge/") + fileName
                    Try
                        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)  
                        fileExtension = fileName.Substring(fileName.IndexOf("."))
                        Dim filenane As String

                        filenane = Now.Year() & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & Now.Millisecond & i
                        fileExtension = filenane & fileExtension

                        Dim destPath As String = Server.MapPath("/public/Data/") + fileExtension
                        System.IO.File.Move(sourcePath, destPath)

                        sqlString = "INSERT INTO KnowledgePicInfo (picTitle, picPath, picStatus, parentIcuitem, picType, picOrder)"
                        sqlString &= "VALUES(@stitle,@route,'W',@parentIcuitem,'A',@picOrder)"

                        myCommand = New SqlCommand(sqlString, myConnection)
                        myCommand.Parameters.Add(New SqlParameter("@stitle", hints(i)))
                        myCommand.Parameters.Add(New SqlParameter("@route", "/public/Data/" & fileExtension))
                        myCommand.Parameters.Add(New SqlParameter("@parentIcuitem", MaxIcuitem))
                        myCommand.Parameters.Add(New SqlParameter("@picOrder", i + 1))

                        myCommand.ExecuteNonQuery()
                        myCommand.Dispose()
                        i = i + 1
                    Catch ex As Exception
                        Response.Write(ex)
                    Finally
                        System.IO.File.Delete(sourcePath)
                        Dim originalFileName = fileName.Replace("_", "")
                        System.IO.File.Delete(Server.MapPath("/knowledge/") + originalFileName)
                    End Try
                End If
            Next

            '---更新討論數---
            sqlString = "UPDATE KnowledgeForum SET DiscussCount = DiscussCount + 1 WHERE gicuitem = @agicuitem"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@agicuitem", Request.QueryString("ArticleId"))
            myCommand.ExecuteNonQuery()
            myCommand.Dispose()
            KFlag = True

            '---20080927---vincent---kpi share---
            Dim share As New KPIShare(MemberId, "shareDiscuss", MaxIcuitem)
            share.HandleShare()

        Catch ex As Exception
            If Request.QueryString("debug") = "true" Then
                Response.Write(ex.ToString)
                Response.End()
            End If
            Response.Redirect("/knowledge/knowledge_cp.aspx?ArticleId=" & Request.QueryString("ArticleId") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    '---活動專用--------------------------------------------------------------------------------------------------
    Dim Activity As New ActivityFilter(WebConfigurationManager.AppSettings("ActivityId"), MemberId, Request.QueryString("ArticleId"))
    Dim ActivityFlag As Boolean = Activity.CheckActivity

    If ActivityFlag Then
		'判斷是否為活動題目
		If Activity.CheckAvtivityKnowledge Then
		  try
			'是的話加分
			sqlString = "INSERT INTO KnowledgeActivity (CUItemId ,Type ,Grade ,MemberId ,State) "
			sqlString &= "VALUES(@CUItemId, 3, 2, @MemberId, 1)"

			myConnection.Open()
			myCommand = New SqlCommand(sqlString, myConnection)
			myCommand.Parameters.AddWithValue("@CUItemId", MaxIcuitem)
			myCommand.Parameters.AddWithValue("@MemberId", MemberId)

			myCommand.ExecuteNonQuery()
			myCommand.Dispose()
		  Catch ex As Exception
            Response.Write(ex)
          End Try
		End If
      'Call Activity.ProcessActivity()
	 
	  
    End If
    'Call Activity.DiscussGradeAdd()
    '-------------------------------------------------------------------------------------------------------------

    'If ActivityFlag Then
      REM SuccessScript = "<script>alert('發表討論成功');window.location.href='/knowledge/knowledge_alp_action.aspx?ArticleId=" & Request.QueryString("ArticleId") & _
              REM "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
      REM Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "submit", SuccessScript)
    'Else
        If KFlag Then

            Dim reDirectURL As String = "/knowledge/knowledge_cp.aspx?ArticleId=" & Request.QueryString("ArticleId") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId
            Dim myArticleId As Integer
            If Integer.TryParse(Request.QueryString("ArticleId"), myArticleId) Then
                '發送通知訊息
                'Modified   By  Leo     2011-07-14  移除通知的功能-----Start------
                'KnowledgeNoticeMessage.SendMessage(myArticleId, MemberId, ArticleContent, reDirectURL)
                'Modified   By  Leo     2011-07-14  移除通知的功能------End-------
            End If


            SuccessScript = "<script>alert('發表討論成功');window.location.href='" & reDirectURL & "';</script>"
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "submit", SuccessScript)
        End If
    'End If

  End Sub

  Protected Sub TempSubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TempSubmitBtn.Click
        Dim oldGuid = ""
        If Not (Request.Form("MemberGuid").ToString() = "") Then
            oldGuid = Request.Form("MemberGuid").ToString()
        End If
    Dim PicInput As Boolean = False
    '---handle picture---
        If (Session(oldGuid & "CaptchaImageText") = Nothing) Then
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼失效!!');window.location.href='" & Request.RawUrl() & "';</script>")
            Exit Sub
        ElseIf (MemberCaptChaTBox.Text <> Session(oldGuid & "CaptchaImageText").ToString()) Then
            ClientScript.RegisterClientScriptBlock(Me.GetType(), "ForgetPassWordScript", "<script language='javascript'>alert('驗證碼錯誤!!');history.back();</script>")
            Exit Sub
        End If
        Session.Remove(oldGuid & "CaptchaImageText")
    Dim MemberId As String = ""
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='" & Request.UrlReferrer.ToString() &"';</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------

    If DiscussContent.Text = "請在這裡輸入您對本問題的討論或想法" Then
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", "<script>alert('請輸入討論內容!!');history.back();</script>")
      Exit Sub
    End If


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
      CategoryId = "A"
    Else
      CategoryId = Request.QueryString("CategoryId")
    End If


    Dim iBaseDSD As String = WebConfigurationManager.AppSettings("KnowledgeBaseDSD")
    Dim iCtUnit As String = WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId")
    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim SuccessScript As String = ""
    Dim MaxIcuitem As String = ""
    Dim ArticleContent As String = DiscussContent.Text.Trim
    ArticleContent = ArticleContent.Replace("<", "&lt;").Replace(">", "&gt;").Replace("%3C", "&lt;").Replace("%3E", "&gt;")

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "INSERT INTO CuDTGeneric (iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, xBody, siteId) "
      sqlString &= "VALUES(@ibasedsd, @ictunit, 'N', @stitle, @ieditor, '0', @xbody, @siteid)"

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ibasedsd", iBaseDSD)
      myCommand.Parameters.AddWithValue("@ictunit", iCtUnit)
      myCommand.Parameters.AddWithValue("@stitle", "討論-" & Request.QueryString("ArticleId"))
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
      myCommand.Parameters.AddWithValue("@parenticuitem", Request.QueryString("ArticleId"))
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

            ' 2010/5/26 vivian modify - COAKM10099043上傳圖片裁切功能(FileName存在HiddenField)
            ' 2010/8/2 vivian modify - 可上傳多張多圖，依"uploadPicCount"決定
            Dim i As Integer = 0
            Dim files(), hints() As String
            files = hidFileName.Value.Split("^")
            hints = hidFileContent.Value.Split("^")

            'Dim myfileuploadcontent As String = hidFileContent.value
            For Each fileName As String In files
                If fileName <> "" Then
                    'If hidFileName.value <> "" Then
                    Dim sourcePath As String = Server.MapPath("/knowledge/") + fileName
                    Try
                        Dim fileExtension As String     '宣告一個變數存放檔案格式(副檔名)  
                        fileExtension = fileName.Substring(fileName.IndexOf("."))
                        Dim filenane As String

                        filenane = Now.Year() & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & Now.Millisecond & i
                        fileExtension = filenane & fileExtension

                        Dim destPath As String = Server.MapPath("/public/Data/") + fileExtension
                        System.IO.File.Move(sourcePath, destPath)

                        sqlString = "INSERT INTO KnowledgePicInfo (picTitle, picPath, picStatus, parentIcuitem, picType, picOrder)"
                        sqlString &= "VALUES(@stitle,@route,'W',@parentIcuitem,'A',@picOrder)"

                        myCommand = New SqlCommand(sqlString, myConnection)
                        myCommand.Parameters.Add(New SqlParameter("@stitle", hints(i)))
                        myCommand.Parameters.Add(New SqlParameter("@route", "/public/Data/" & fileExtension))
                        myCommand.Parameters.Add(New SqlParameter("@parentIcuitem", MaxIcuitem))
                        myCommand.Parameters.Add(New SqlParameter("@picOrder", i + 1))

                        myCommand.ExecuteNonQuery()
                        myCommand.Dispose()
                        i = i + 1
                    Catch ex As Exception
                        Response.Write(ex)
                    Finally
                        System.IO.File.Delete(sourcePath)
                        Dim originalFileName = fileName.Replace("_", "")
                        System.IO.File.Delete(Server.MapPath("/knowledge/") + originalFileName)
                    End Try
                End If
            Next

            SuccessScript = "<script>alert('發表暫存討論成功');window.location.href='/knowledge/knowledge_cp.aspx?ArticleId=" & Request.QueryString("ArticleId") & _
            "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "';</script>"
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "submit", SuccessScript)

        Catch ex As Exception

            If Request.QueryString("debug") = "true" Then
                Response.Write(ex.ToString)
                Response.End()
            End If
            Response.Redirect("/knowledge/knowledge_cp.aspx?ArticleId=" & Request.QueryString("ArticleId") & "&ArticleType=" & Request.QueryString("ArticleType") & "&CategoryId=" & Request.QueryString("CategoryId"))
            Response.End()

        Finally

            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If

        End Try

  End Sub

  Protected Sub ResetBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ResetBtn.Click

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
      CategoryId = "A"
    Else
      CategoryId = Request.QueryString("CategoryId")
    End If

    Response.Redirect("/knowledge/knowledge_cp.aspx?ArticleId=" & Request.QueryString("ArticleId") & "&ArticleType=" & ArticleType & "&CategoryId=" & CategoryId)

  End Sub

    Protected Sub deleteBtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles deleteBtn.Click
        If hidFileName.Value <> "" And hidFileName.Value <> "undefined" Then
            Dim sourcePath As String = Server.MapPath("/knowledge/") + hidFileName.value
            System.IO.File.Delete(sourcePath)
            Dim originalFileName = hidFileName.value.Replace("_", "")
            System.IO.File.Delete(Server.MapPath("/knowledge/") + originalFileName)
            hidFileName.Value = ""
            hidFileContent.Value = ""
            lblImgHint.Text = ""
            previewImg.Visible = False
            uploadBtn.Visible = True
            deleteBtn.Visible = False
        End If
    End Sub

    Function GetPicInfo(ByVal aid As String, ByVal maxWidth As Integer) As String

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
                str &= "<div style='width:" & maxWidth & "px'>"
                str &= "<a href=""" & myReader("picPath") & """ alt=""" & myReader("picTitle") & """ target='_blank'><img src=""" & myReader("picPath") & """ alt=""" & myReader("picTitle") & """ style='max-width:" & maxWidth & "px' /></a>"
                str &= "<p style='color:#666666'>" & myReader("picTitle") & "</p>"
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
            myConnection.Close()
        End Try

        Return str

    End Function
    'Modify by Max 2011/06/07
    Function ReplaceAndFindKeyword(ByVal strTemp As String) As String
        strTemp = strTemp.replace(Chr(13), "<br />")
        Dim string1
        string1 = PediaUtility.ReplacePedia(strTemp)
        strTemp = string1(0)
        'innerTag &= string1(1)

        Return strTemp
    End Function
    'End
End Class
