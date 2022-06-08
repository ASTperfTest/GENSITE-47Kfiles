Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_myknowledge_changepic
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim sb As New StringBuilder
    Dim MemberId As String
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>"

    '-----------------------------------------------------------------------------
    If Session("memID") = "" Or Session("memID") Is Nothing Then
      MemberId = ""
      Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
      Exit Sub
    Else
      MemberId = Session("memID").ToString()
    End If
    '-----------------------------------------------------------------------------
    TabText.Text = "<ul><li class=""current""><a href=""/knowledge/myknowledge_record.aspx""><span>我的紀錄</span></a></li>" & _
                   "<li><a href=""/knowledge/myknowledge_question_lp.aspx""><span>我的發問</span></a></li>" & _
                   "<li><a href=""/knowledge/myknowledge_discuss_lp.aspx""><span>我的討論</span></a></li>" & _
                   "<li><a href=""/knowledge/myknowledge_trace_lp.aspx""><span>我的追蹤</span></a></li>" & _
                   "<li><a href=""/knowledge/myknowledge_pedia.aspx""><span>我的小百科</span></a></li></ul>"
    '-----------------------------------------------------------------------------

    memberIdText.Text = MemberId

    If Not Page.IsPostBack Then

      Dim picpath As String = GetMemberPicType(MemberId)
      
      If picpath = "A" Then
        picTypeA.Checked = True
      End If
      If picpath = "B" Then
        picTypeB.Checked = True
      End If
      If picpath = "C" Then
        picTypeC.Checked = True
      End If
      If picpath = "D" Then
        picTypeD.Checked = True
      End If

    End If

  End Sub

  Function GetMemberPicType(ByVal memberId As String) As String

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlStr As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    myConnection = New SqlConnection(ConnString)
    Try
      sqlStr = "SELECT pictype FROM Member WHERE account = @memberid"
      myConnection.Open()
      myCommand = New SqlCommand(sqlStr, myConnection)
      myCommand.Parameters.AddWithValue("@memberid", memberId)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        If myReader("picpath") IsNot DBNull.Value Then
          Return myReader("picpath").ToString()
        Else
          Return "A"
        End If
      Else
        Return "A"
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()
    Catch ex As Exception
      Return "A"
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Function

  Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SubmitBtn.Click

    Dim type As String = ""
    Dim script As String = "<script>alert('修改成功!!');window.location.href='/knowledge/myknowledge_record.aspx';</script>"

    If picTypeA.Checked Then type = "A"
    If picTypeB.Checked Then type = "B"
    If picTypeC.Checked Then type = "C"
    If picTypeD.Checked Then type = "D"

    SaveMemberPicPath(Session("memID").ToString(), type)

    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "done", script)

  End Sub

  Protected Sub CancelBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelBtn.Click

    Response.Redirect("/knowledge/myknowledge_record.aspx")

  End Sub

  Sub SaveMemberPicPath(ByVal memberId As String, ByVal pictype As String)

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sqlStr As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      sqlStr = "UPDATE Member SET pictype = @pictype WHERE account = @account"
      myCommand = New SqlCommand(sqlStr, myConnection)
      myCommand.Parameters.AddWithValue("@pictype", pictype)
      myCommand.Parameters.AddWithValue("@account", memberId)
      myCommand.ExecuteNonQuery()
      myCommand.Dispose()
    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

End Class
