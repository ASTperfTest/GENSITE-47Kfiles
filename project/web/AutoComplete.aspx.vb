Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration


Partial Class AutoComplete
    Inherits System.Web.UI.Page

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ODBCDSN").ConnectionString
    Dim sqlString As String = ""

    Dim myReader As SqlDataReader
    Dim myTable As DataTable
    Dim myDataRow As DataRow

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        getAutoCompleteData()
    End Sub

    Private Sub getAutoCompleteData()
        Dim prefixText As String = Context.Request.QueryString(0)
        Response.Clear()
        Response.Charset = "utf-8"
        Response.Buffer = True
        EnableViewState = False
        Response.ContentEncoding = System.Text.Encoding.UTF8
        Response.ContentType = "text/plain"

        Using myConnection As New SqlConnection(ConnString)
            Try
                sqlString = "SELECT CatTreeRoot.CtRootName "
                sqlString &= "FROM NodeInfo  "
                sqlString &= "INNER JOIN CatTreeRoot ON CatTreeRoot.CtRootID = NodeInfo.CtrootID  "
                sqlString &= "INNER JOIN [Type] ON NodeInfo.type1 = Type.classname  "
                sqlString &= "WHERE (CatTreeRoot.inUse = 'Y')  "
                If prefixText <> "" Then
                    sqlString &= "and (CatTreeRoot.CtRootName like '%" + prefixText + "%') "
                End If
                sqlString &= "order by  "
                sqlString &= "	(Case When NodeInfo.order_num is null Then 1 Else 0 End) "
                sqlString &= "	, order_num DESC "
                sqlString &= "	, CatTreeRoot.CtrootID "

                myConnection.Open()
                Dim myCommand As SqlCommand = New SqlCommand(sqlString, myConnection)
                myReader = myCommand.ExecuteReader()

                If myReader.HasRows Then
                    While myReader.Read

                        Dim item As String = myReader.Item(0) + vbNewLine
                        Response.Write(item)
                    End While
                End If

            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            End Try
        End Using
    End Sub

End Class