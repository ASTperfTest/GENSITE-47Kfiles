Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration


Partial Class Pedia_PediaExplainContent
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


    Dim aid As String = Request.QueryString("AId")
    Dim paid As String = Request.QueryString("PAId")

    '---kpi use---reflash wont add grade---
    If Request.QueryString("kpi") <> "0" Then
      '---start of kpi user---20080911---vincent---
      If Session("memID") IsNot Nothing Or Session("memID") <> "" Then
        Dim browse As New KPIBrowse(Session("memID"), "browsePediaExplainCP", aid)
        browse.HandleBrowse()
      End If
      '---end of kpi use---
      Dim relink As String = "/Pedia/PediaExplainContent.aspx?AId=" & aid & "&PAId=" & paid & "&kpi=0"
      Response.Redirect(relink)
      Response.End()
    End If
    '---end of kpi use---  

    AddLink.NavigateUrl = "/Pedia/PediaExplainAdd.aspx?AId=" & aid & "&PAId=" & paid
    BackToContent.NavigateUrl = "/Pedia/PediaContent.aspx?AId=" & paid

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT CuDTGeneric.sTitle, Pedia.engTitle, CuDTGeneric.vGroup FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (CuDTGeneric.iCUItem = @icuitem)"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", paid)
      myReader = myCommand.ExecuteReader

      If myReader.Read Then

        WordTitle.Text = "<a href=""javascript:WordSearch('" & myReader("sTitle") & "')"">" & myReader("sTitle") & "</a>"
        WordEngTitle.Text = IIf(myReader("engTitle") Is DBNull.Value Or myReader("engTitle") = "", "&nbsp;", myReader("engTitle"))

        If myReader("vGroup") IsNot DBNull.Value Then
          If myReader("vGroup").ToString = "Y" Then
            AddLink.Visible = True
          Else
            AddLink.Visible = False
          End If
        Else
          AddLink.Visible = False
        End If

      Else
        Session("PediaErrMsg") = "詞目不存在"
        ShowErrorMessage()
        Exit Sub
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "詞目錯誤"
      ShowErrorMessage()
      Exit Sub
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "  SELECT CuDTGeneric.xBody, Member.realname, Member.nickname, Pedia.commendTime FROM CuDTGeneric INNER JOIN "
      sql &= " Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem INNER JOIN Member ON Pedia.memberId = Member.account "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (CuDTGeneric.iCUItem = @icuitem)"

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", aid)
      myReader = myCommand.ExecuteReader

      If myReader.Read Then

        If myReader("nickname") IsNot DBNull.Value Then
          If myReader("nickname") <> "" Then
            MemberName.Text = myReader("nickname")
          Else
            MemberName.Text = myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1)
          End If
        ElseIf myReader("realname") IsNot DBNull.Value Then
          If myReader("realname") <> "" Then
            MemberName.Text = myReader("realname").ToString.Substring(0, 1) & "＊" & myReader("realname").ToString.Substring(myReader("realname").ToString.Trim.Length - 1, 1)
          Else
            MemberName.Text = "&nbsp;"
          End If
        Else
          MemberName.Text = "&nbsp;"
        End If

        ExplainDate.Text = Date.Parse(myReader("commendTime")).ToShortDateString
        ExplainBody.Text = "<p>" & myReader("xBody").ToString.Replace(vbCrLf, "<br />") & "</p>"

      Else
        Session("PediaErrMsg") = "補充內容不存在"
        ShowErrorMessage()
        Exit Sub
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

    Catch ex As Exception
      Session("PediaErrMsg") = "補充內容錯誤"
      ShowErrorMessage()
      Exit Sub
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try
        '把count移到viewcounter.aspx 
        'myConnection = New SqlConnection(ConnString)
        'Try
        '  Sql = " UPDATE CuDTGeneric SET ClickCount = ClickCount + 1 WHERE (CuDTGeneric.iCUItem = @icuitem)"

        '  myConnection.Open()
        '  myCommand = New SqlCommand(Sql, myConnection)
        '  myCommand.Parameters.AddWithValue("@icuitem", aid)
        '  myCommand.ExecuteNonQuery()
        '  myCommand.Dispose()

        'Catch ex As Exception
        'Finally
        '  If myConnection.State = ConnectionState.Open Then
        '    myConnection.Close()
        '  End If
        'End Try

  End Sub

  Sub ShowErrorMessage()

    Dim str = "<script>alert('" & Session("PediaErrMsg") & "');window.location.href='/Pedia/PediaList.aspx';</script>"

    Session("PediaErrMsg") = ""

    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "PediaError", str)

  End Sub


End Class
