Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Public Class ArticleSearch
    Inherits System.Web.UI.Page

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ODBCDSN").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            getddlKindData()
        End If
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReset.Click
        ddlKind.SelectedIndex = 0
        txtSubject.Text = ""
        txtTitle.Text = ""
        txtContent.Text = ""
        txtFrom.Text = ""
        txtTo.Text = ""
        'getddlSubjectData()
    End Sub

    Private Sub getddlKindData()
        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT classid, classname FROM Type WHERE (datalevel = 1) "
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myReader = myCommand.ExecuteReader
            ddlKind.Items.Clear()
            ddlKind.Items.Add(New ListItem("全部", "0"))
            While myReader.Read
                ddlKind.Items.Add(New ListItem(myReader.Item("classname").ToString(), myReader.Item("classid").ToString()))
            End While
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
        Catch ex As Exception
            If Request("debug") = "true" Then
                Response.Write(ex.ToString())
                Response.End()
            End If
            Response.Redirect("/mp.asp?mp=1")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try
        'getddlSubjectData()
    End Sub

    'Private Sub getddlSubjectData()
    '    myConnection = New SqlConnection(ConnString)
    '    Try
    '        sqlString = "SELECT "
    '        sqlString += "	 NodeInfo.CtrootID "
    '        sqlString += "	,CatTreeRoot.CtRootName "
    '        sqlString += "from NodeInfo "
    '        sqlString += "inner join CatTreeRoot on CatTreeRoot.CtRootID = NodeInfo.CtrootID "
    '        sqlString += "where (CatTreeRoot.inUse = 'Y') "
    '        If ddlKind.SelectedValue <> "0" Then
    '            sqlString += "and NodeInfo.type1 = '" + ddlKind.SelectedItem.ToString() + "'"
    '        End If
    '        myConnection.Open()
    '        myCommand = New SqlCommand(sqlString, myConnection)
    '        myReader = myCommand.ExecuteReader
    '        ddlSubject.Items.Clear()
    '        ddlSubject.Items.Add(New ListItem("全部", "0"))
    '        While myReader.Read
    '            ddlSubject.Items.Add(New ListItem(myReader.Item("CtRootName").ToString(), myReader.Item("CtrootID").ToString()))
    '        End While
    '    Catch ex As Exception
    '        If Request("debug") = "true" Then
    '            Response.Write(ex.ToString())
    '            Response.End()
    '        End If
    '        Response.Redirect("/mp.asp?mp=1")
    '        Response.End()
    '    Finally
    '        If myConnection.State = ConnectionState.Open Then
    '            myConnection.Close()
    '        End If
    '    End Try
    'End Sub

    'Protected Sub ddlKind_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlKind.SelectedIndexChanged
    '    getddlSubjectData()
    'End Sub

    Protected Sub btnSearch_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSearch.Click
        'Dim pSubjectId As String = ddlSubject.SelectedValue
        Dim pSubjectName As String = Server.UrlEncode(txtSubject.Text.Trim())
        Dim pType As String = ddlKind.SelectedItem.ToString()
        Dim pTitle As String = Server.UrlEncode(txtTitle.Text.Trim())
        Dim pBody As String = Server.UrlEncode(txtContent.Text.Trim())
        Dim pPostDate As String
        If (txtFrom.Text.Trim <> "" And txtTo.Text.Trim <> "") Then
            pPostDate = Convert.ToDateTime(txtFrom.Text.Trim()).ToString("yyyyMMdd") + Convert.ToDateTime(txtTo.Text.Trim()).ToString("yyyyMMdd")
        End If
        If pType = "全部" Then
            pType = ""
        End If
        Response.Redirect("~/ArticleResult.aspx?SubjectName=" + pSubjectName + "&Title=" + pTitle + "&Body=" + pBody + "&PDate=" + pPostDate + "&Type=" + pType)
    End Sub
End Class