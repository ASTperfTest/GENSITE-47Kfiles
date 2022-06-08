Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_knowledgeExpertReply
    Inherits System.Web.UI.Page

  Dim connString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
  Dim sql As String = ""
  Dim myConnection As SqlConnection
  Dim myCommand As SqlCommand
  Dim myReader As SqlDataReader
  Dim articleId As String
  Dim DArticleId As String
  Dim articleRand As String
  Dim expertId As String

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    articleId = Request.QueryString("ArticleId")
    DArticleId = Request.QueryString("DArticleId")
    articleRand = Request.QueryString("rand")
    expertId = Request.QueryString("expertId")

    If articleId = "" Or DArticleId = "" Or articleRand = "" Or expertId = "" Then
      Response.Redirect("/knowledge/knowledge.aspx")
      Response.End()
    End If

    '---檢查連結---
    Dim flag As Boolean = True
    If flag Then
      flag = CheckReply()
      If Not flag Then RedirectToMainPage()
    End If
    If flag Then
      '---檢查是否回覆過---
      flag = CheckHaveReply()
      If Not flag Then AlertAndRedirectToMainPage("您已回覆過此篇文章!!")
    End If



    If Not Page.IsPostBack Then
      If flag Then
        Try
          myConnection.Open()
          sql = "SELECT * FROM Member WHERE account = @account"
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@account", expertId)
          myReader = myCommand.ExecuteReader
          If myReader.Read Then
            ExpertName.Text = myReader("account").ToString.Trim & "&nbsp;" & myReader("realname").ToString.Trim
          Else
            flag = False
          End If
          If Not myReader.IsClosed Then myReader.Close()
        Catch ex As Exception
        Finally
          If myConnection.State = ConnectionState.Open Then myConnection.Close()
        End Try
        If Not flag Then RedirectToMainPage()
      End If
      If flag Then
        Try
          myConnection.Open()
          sql = "SELECT icuitem, topCat, sTitle, ISNULL(xBody, '') as xBody FROM CuDTGeneric "
          sql &= "INNER JOIN KnowledgeForum ON CuDTGeneric.icuitem = KnowledgeForum.gicuitem "
          sql &= "WHERE CuDTGeneric.icuitem = @articleid"
          myCommand = New SqlCommand(sql, myConnection)
          myCommand.Parameters.AddWithValue("@articleid", articleId)
          myReader = myCommand.ExecuteReader
          If myReader.Read Then
            QuestionTitle.Text = myReader("sTitle")
            QuestionContent.Text = myReader("xBody")
            QuestionLink.NavigateUrl = "/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("icuitem") & "&ArticleType=A&CategoryId=" & myReader("topCat")
            QuestionLink.Text = WebConfigurationManager.AppSettings("WWWUrl") & "/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("icuitem") & "&ArticleType=A&CategoryId=" & myReader("topCat")
          Else
            flag = False
          End If
          If Not myReader.IsClosed Then myReader.Close()
        Catch ex As Exception
        Finally
          If myConnection.State = ConnectionState.Open Then myConnection.Close()
        End Try
        If Not flag Then RedirectToMainPage()
      End If
    End If
  End Sub

  Function CheckReply() As Boolean

    Dim flag As Boolean = True
    myConnection = New SqlConnection(connString)
    Try
      myConnection.Open()
      sql = "SELECT * FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.icuitem = KnowledgeForum.gicuitem "
      sql &= "WHERE CuDTGeneric.icuitem = @darticleid"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@darticleid", DArticleId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("ParentIcuitem") <> articleId Then flag = False
        If myReader("expertRand") <> articleRand Then flag = False
        If myReader("expertId") <> expertId Then flag = False
      Else
        flag = False
      End If
      If Not myReader.IsClosed Then myReader.Close()
    Catch ex As Exception
      flag = False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return flag

  End Function

  Function CheckHaveReply() As Boolean

    Dim flag As Boolean = True
    myConnection = New SqlConnection(connString)
    Try
      myConnection.Open()
      sql = "SELECT * FROM KnowledgeForum WHERE gicuitem = @darticleid"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@darticleid", DArticleId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("expertReply") = "Y" Then flag = False      
      End If
      If Not myReader.IsClosed Then myReader.Close()
    Catch ex As Exception      
      flag = False
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    Return flag

  End Function

  Sub RedirectToMainPage()
    Response.Redirect("/knowledge/knowledge.aspx")
    Response.End()
  End Sub

  Sub AlertAndRedirectToMainPage(ByVal msg As String)

    Dim str As String = "<script>alert('" & msg & "');window.location.href='/knowledge/knowledge.aspx';</script>"
    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "havereply", str)
    
  End Sub

  Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles SubmitBtn.Click

    Dim script As String = "<script language=""javascript"">alert('專家補充完成');window.location.href='" & QuestionLink.NavigateUrl & "';</script>"

    Dim xbody As String = ExpertAddText.Text
    xbody = xbody.Replace("<", "&lt;").Replace(">", "&gt;")

    Dim connString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand

    myConnection = New SqlConnection(connString)
    Try
      myConnection.Open()
      sql = "UPDATE CuDTGeneric SET xPostDate = GETDATE(), xBody = @xbody WHERE icuitem = @icuitem"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", DArticleId)
      myCommand.Parameters.AddWithValue("@xbody", xbody)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sql = "UPDATE KnowledgeForum SET Status = 'N', expertReply = 'Y', expertReplyDate = GETDATE() WHERE gicuitem = @gicuitem"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@gicuitem", DArticleId)      
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      sql = "UPDATE KnowledgeForum SET HavePros = 'Y' WHERE gicuitem = @gicuitem"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@gicuitem", articleId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()

      ClientScript.RegisterClientScriptBlock(Me.GetType, "success", script)

    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try

  End Sub

  Protected Sub ResetBtn_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ResetBtn.Click

    Dim url As String = "/knowledge/knowledgeExpertReply.aspx?ArticleId={0}&DArticleId={1}&rand={2}&expertId={3}"
    url = url.Replace("{0}", Request.QueryString("ArticleId"))
    url = url.Replace("{1}", Request.QueryString("DArticleId"))
    url = url.Replace("{2}", Request.QueryString("rand"))
    url = url.Replace("{3}", Request.QueryString("expertId"))

    Response.Redirect(url)
    Response.End()

  End Sub

End Class
