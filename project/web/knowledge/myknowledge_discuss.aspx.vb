Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class myknowledge_discuss
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ArticleId As String
    Dim MemberId As String
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>"
    Dim Script1 As String = "<script>alert('連結錯誤!!');window.location.href='/knowledge/myknowledge_discuss_lp.aspx';</script>"
    Dim ScriptFail As String = "<script>alert('系統錯誤!!');window.location.href='/knowledge/myknowledge_discuss_lp.aspx';</script>"
    Dim sb As StringBuilder
    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim index As Integer = 0

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

    Dim ArticleType As String
    Dim CategoryId As String
    Dim Keyword As String
    Dim Sort As String

    If Request.QueryString("ArticleType") <> "A" And Request.QueryString("ArticleType") <> "B" And Request.QueryString("ArticleType") <> "C" _
                And Request.QueryString("ArticleType") <> "D" And Request.QueryString("ArticleType") <> "E" And Request.QueryString("ArticleType") <> "F" Then
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

    If Request.QueryString("Sort") <> "A" And Request.QueryString("Sort") <> "D" Then
      Sort = "D"
    Else
      Sort = Request.QueryString("Sort")
    End If

    If Request.QueryString("keyword") IsNot Nothing Then
      Keyword = Web.HttpUtility.UrlDecode(Request.QueryString("keyword"))
    Else
      Keyword = ""
    End If

    GoBackLink.NavigateUrl = "/knowledge/myknowledge_discuss_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()

      Dim Atype As String = ""
      Dim DArticleId As String = ""
      Dim Body As String = ""
	  Dim PostTime As String = ""

      Atype = Request("Atype")
      DArticleId = Request("DArticleId")
      Body = Request("ArticleContent")
      PostTime = Request("PostTime")

      If Atype <> "" And DArticleId <> "" And Body <> "" And PostTime <> "" Then

        '---更新狀態---

        If Atype = "0" Then
          sqlString = "UPDATE CuDTGeneric SET dEditDate = GETDATE(), xBody = @xBody WHERE IcuItem = @agicuitem"
        ElseIf Atype = "1" Then
          sqlString = "UPDATE CuDTGeneric SET fCTUPublic = 'Y', xPostDate = GETDATE(), xBody = @xBody WHERE IcuItem = @agicuitem"
        End If

        myCommand = New SqlCommand(sqlString, myConnection)
        myCommand.Parameters.AddWithValue("@xBody", Body)
        myCommand.Parameters.AddWithValue("@agicuitem", DArticleId)
        myCommand.ExecuteNonQuery()
        myCommand.Dispose()

        If Atype = "1" Then
          '---更新討論數---
          sqlString = "UPDATE KnowledgeForum SET DiscussCount = DiscussCount + 1 WHERE gicuitem = @agicuitem"
          myCommand = New SqlCommand(sqlString, myConnection)
          myCommand.Parameters.AddWithValue("@agicuitem", ArticleId)
          myCommand.ExecuteNonQuery()
          myCommand.Dispose()

          '---活動專用--------------------------------------------------------------------------------------------------
          Dim Activity As New ActivityFilter(WebConfigurationManager.AppSettings("ActivityId"), MemberId, ArticleId)
          Dim ActivityFlag As Boolean = Activity.CheckActivity

          If ActivityFlag Then
            'Call Activity.ProcessActivity()
			'判斷是否為活動題目
			If Activity.CheckAvtivityKnowledge Then
			  try
				'是的話加分
				'活動前的討論得1分,活動中的得2分
				Dim grade As String = ""
				sqlString = "SELECT ActivityId FROM Activity WHERE ( CONVERT(DATETIME,@PostTime) BETWEEN ActivityStartTime AND ActivityEndTime) AND ActivityId =@ActivityId"

				myCommand = New SqlCommand(sqlString, myConnection)
				myCommand.Parameters.AddWithValue("@PostTime", PostTime)
				myCommand.Parameters.AddWithValue("@ActivityId", WebConfigurationManager.AppSettings("ActivityId"))
				myReader = myCommand.ExecuteReader()
				'Response.Write(PostTime)
				'Response.Write("<br/>" & sqlString)
				'Response.Write("<br/>" & WebConfigurationManager.AppSettings("ActivityId"))
				If myReader.Read Then
					grade = "2"
				Else
					grade = "1"
				End If
				
				If Not myReader.IsClosed Then
				myReader.Close()
				End If
				'Response.Write("<br/>" & grade)

				myCommand.Dispose()
	  
				'新增活動積分
				sqlString = "INSERT INTO KnowledgeActivity (CUItemId ,Type ,Grade ,MemberId ,State) "
				sqlString &= "VALUES(@CUItemId, 3, @Grade, @MemberId, 1)"
			
				myCommand = New SqlCommand(sqlString, myConnection)
				myCommand.Parameters.AddWithValue("@CUItemId", DArticleId)
				myCommand.Parameters.AddWithValue("@Grade", grade)
				myCommand.Parameters.AddWithValue("@MemberId", MemberId)

				myCommand.ExecuteNonQuery()
				myCommand.Dispose()
			  Catch ex As Exception
				Response.Write(ex)
			  End Try
			End If
          End If
          'Call Activity.DiscussGradeAdd()
          '-------------------------------------------------------------------------------------------------------------

          '---20080927---vincent---kpi share---
          Dim share As New KPIShare(MemberId, "shareDiscuss", "")
          share.HandleShare()

                End If

                If Atype = "0" Then
                    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "success", "<script>alert('討論暫存成功!!');</script>")
                ElseIf Atype = "1" Then
                    '發送通知訊息
                    'Modified   By  Leo     2011-07-14  移除通知的功能-----Start------
                    'KnowledgeNoticeMessage.SendMessage(Integer.Parse(Request.QueryString("ArticleId")), MemberId, Body, "")
                    'Modified   By  Leo     2011-07-14  移除通知的功能------End-------

                    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "success", "<script>alert('討論發表成功!!');window.location.href='" & Request.RawUrl() & "';</script>")
                End If

            End If


      sqlString = "SELECT CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CodeMain.mValue, CuDTGeneric.xBody FROM CuDTGeneric "
      sqlString &= "INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sqlString &= "WHERE (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.iCUItem = @icuitem) AND (CuDTGeneric.fCTUPublic = 'Y') AND (KnowledgeForum.Status = 'N') "

      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
      myReader = myCommand.ExecuteReader()

      If myReader.Read Then
        QuestionTitleText.Text = myReader("sTitle")
        QuestionPostDateText.Text = Date.Parse(myReader("xPostDate")).ToShortDateString
        QuestionCategoryText.Text = myReader("mValue")
        QuestionBodyText.Text = myReader("xBody").replace(Chr(13), "<br />")
      Else
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NotDound", "<script>alert('此問題不公開，無法編輯討論。');history.back();</script>")
        Exit Sub
      End If

      If Not myReader.IsClosed Then
        myReader.Close()
      End If

      myCommand.Dispose()
	  
	  'mod by Ivy 2010.5.11 知識家活動 新增撈取Abstract欄位
      sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, KnowledgeForum.CommandCount, KnowledgeForum.BrowseCount, "
      sqlString &= "KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, CuDTGeneric.fCTUPublic, CuDTGeneric.dEditDate, CuDTGeneric.Abstract "
      sqlString &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sqlString &= "WHERE (KnowledgeForum.ParentIcuitem = @icuitem) AND (CuDTGeneric.iCTUnit = @ictunit) AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iEditor = @ieditor)"

      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
      myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
      myCommand.Parameters.AddWithValue("@ieditor", MemberId)
      myReader = myCommand.ExecuteReader()

      sb = New StringBuilder
      While myReader.Read
        index += 1
        sb.Append("<div class=""mydis"">")
        sb.Append("<label for=""textarea""><span class=""title2"">我的討論(" & index & ")</span> ")

        If myReader("fCTUPublic").ToString() = "Y" Then
          sb.Append("發表於 " & Date.Parse(myReader("xPostDate")).ToShortDateString & "</label>")
        ElseIf myReader("fCTUPublic").ToString() = "N" Then
          sb.Append("草稿存於 " & Date.Parse(myReader("dEditDate")).ToShortDateString & "</label>")
        End If

        sb.Append("<p><textarea id=""Content" & index & """ name=""Content" & index & """>" & myReader("xBody") & "</textarea></p>")
        sb.Append("</div>")

        sb.Append("<table class=""table1"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""排版用表格"">")
        sb.Append("<tr>")
        sb.Append("<td>瀏覽：<span class=""number2"">" & myReader("BrowseCount") & "</span>次</td>")
        sb.Append("<td>意見：<span class=""number2"">" & myReader("CommandCount") & "</span>次</td>")
        sb.Append("<td colspan=""2"">平均評價：")
        sb.Append(ConcatStarString(myReader("GradeCount"), myReader("GradePersonCount")))
        sb.Append("</td>")
        sb.AppendLine("</tr>")
        sb.Append("</table>")

        If myReader("fCTUPublic") = "N" and myReader("Abstract").Tostring() <> "Y" Then

          sb.Append("<div class=""float"">")
          sb.Append("<input class=""btn2"" type=""button"" name=""post"" value=""發布本筆討論"" onclick=""FormSubmit( '1', '" & myReader("iCUItem") & "','" & Date.Parse(myReader("xPostDate")).ToString("yyyy-MM-dd HH:mm:ss") & "', 'Content" & index & "')"" />")
          sb.Append("</div>")

          sb.Append("<div class=""float"">")
          sb.Append("<input class=""btn2"" type=""button"" name=""post"" value=""儲存本筆討論"" onclick=""FormSubmit( '0', '" & myReader("iCUItem") & "','" & Date.Parse(myReader("xPostDate")).ToString("yyyy-MM-dd HH:mm:ss") & "', 'Content" & index & "')"" />")
          sb.Append("</div>")

        End If

        sb.Append("<div class=""space""></div><hr size=""1"" color=""#CCCCCC"" />")

      End While

      QuestionDiscussText.Text = sb.ToString()
      sb = Nothing

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


  Protected Sub BackToDiscussListBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackToDiscussListBtn.Click

    Dim ArticleType As String = Request.QueryString("ArticleType")
    Dim CategoryId As String = Request.QueryString("CategoryId")
    Dim Keyword As String = Web.HttpUtility.UrlDecode(Request.QueryString("Keyword"))
    Dim Sort As String = Request.QueryString("Sort")

    Response.Redirect("/knowledge/myknowledge_discuss_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&Sort=" & Sort)

  End Sub

End Class

