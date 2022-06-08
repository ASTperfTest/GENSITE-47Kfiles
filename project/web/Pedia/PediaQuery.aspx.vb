Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Pedia_PediaQuery
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    If Not Page.IsPostBack Then

      Session("PediaQueryTitle") = ""
      Session("PediaQueryEngTitle") = ""
      Session("PediaQueryBody") = ""
      Session("PediaQueryIsOpen") = ""

    End If

  End Sub

  Protected Sub SubmitBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SubmitBtn.Click

    If WordTitle.Text = "請輸入關鍵字" Then
      WordTitle.Text = ""
    End If
    If WordEngTitle.Text = "請輸入關鍵字" Then
      WordEngTitle.Text = ""
    End If
    If WordBody.Text = "請輸入關鍵字" Then
      WordBody.Text = ""
    End If

    Session("PediaQueryTitle") = WordTitle.Text.Replace("'", "''").Trim
    Session("PediaQueryEngTitle") = WordEngTitle.Text.Replace("'", "''").Trim
    Session("PediaQueryBody") = WordBody.Text.Replace("'", "''").Trim
    Session("PediaQueryIsOpen") = IsOpenDDL.SelectedValue

    Response.Redirect("/Pedia/PediaList.aspx?type=query")

  End Sub

  Protected Sub ResetBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ResetBtn.Click

    Response.Redirect("/Pedia/PediaQuery.aspx")

  End Sub


End Class
