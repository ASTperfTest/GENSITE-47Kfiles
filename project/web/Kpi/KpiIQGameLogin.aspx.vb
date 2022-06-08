Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Kpi_KpiIQGameLogin
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


    Dim account As String = Request.QueryString("account")
    Dim password As String = Request.QueryString("password")
    Dim flag As Boolean = True

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      sql = "SELECT TOP 1 * FROM Member WHERE account = @account AND passwd = @password"
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@account", account)
      myCommand.Parameters.AddWithValue("@password", password)
      myReader = myCommand.ExecuteReader
      If Not myReader.Read Then
        flag = False
      End If

    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

    If flag Then

      Dim login As New KPILogin(account, "loginIQCount", "")
      login.HandleLogin()

    End If
	if(WebUtility.checkParam(account) or WebUtility.checkParam(password) ) then
		Response.Redirect("/")
	End if
	if account = "" then
		Response.Redirect("/")
	End if
    Response.Redirect("/IQGame/checklogin.asp?userIdno=" & account & "&userEmail=" & password & "&kpi=0")

  End Sub

End Class
