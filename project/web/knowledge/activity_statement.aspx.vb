Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class activity_statement
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Dim Activity As New ActivityFilter(System.Web.Configuration.WebConfigurationManager.AppSettings("ActivityId"), "")
    Dim ActivityFlag As Boolean = Activity.CheckActivity
	Dim sb As New StringBuilder
		
	If  ActivityFlag Then
		sb.Append("<a href=""/knowledge/knowledge_alp.aspx""><img src=""image/mn_5.gif"" width=""107"" height=""28"" /></a>")
	Else
		sb.Append("&nbsp;")
	End If
	
	TabText.Text = sb.tostring()
  End Sub
  
End Class
