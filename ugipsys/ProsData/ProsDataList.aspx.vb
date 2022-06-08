Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Web
Imports System.Web.UI.WebControls


Partial Class ProsData_ProsDataList
    Inherits System.Web.UI.Page

    Dim TotalCount As Integer
    Dim conn As SqlConnection
    Dim comm As SqlCommand
    Dim reader As SqlDataReader
    Dim ConnString As String = ConfigurationManager.ConnectionStrings("ConnString").ConnectionString
    Dim sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SqlDataSource1.ConnectionString = ConnString


        sql = "SELECT * FROM Member WHERE actor = '5' ORDER BY createtime DESC"
        SqlDataSource1.SelectCommand = sql
        ListTable.DataBind()

        If Not Page.IsPostBack Then

            Call GetRecordCount()
            Call SetPage()
            Call BackNextEnable()

        End If

    End Sub

    Public Sub GetRecordCount()

        conn = New SqlConnection(ConnString)
        Try
            conn.Open()
            sql = "SELECT COUNT(*) FROM Member WHERE actor = '5'"
            comm = New SqlCommand(sql, conn)
            TotalCount = comm.ExecuteScalar
            TotalCountTxt.Text = TotalCount
            comm.Dispose()

        Catch ex As Exception
        Finally
            If conn.State = ConnectionState.Open Then
                conn.Close()
            End If
        End Try

    End Sub

    Public Sub SetPage()

        PageNumberDDL.Items.Clear()
        Dim Number As Integer, i As Integer

        If TotalCount Mod ListTable.PageSize = 0 Then
            Number = TotalCount / ListTable.PageSize
        Else
            If TotalCount < PageSizeDDL.SelectedValue Then
                Number = (TotalCount / ListTable.PageSize)
            Else
                Number = (TotalCount / ListTable.PageSize) + 1
            End If

        End If
        For i = 0 To Number
            Dim LItem As New ListItem(Convert.ToString(i + 1), Convert.ToString(i + 1))
            PageNumberDDL.Items.Add(LItem)
            LItem = Nothing
        Next

    End Sub

    Public Sub BackNextEnable()

        If ListTable.PageIndex = 0 Then
            BackLBtn.Enabled = False
            If ListTable.PageCount > 1 Then
                NextLBtn.Enabled = True
            Else
                NextLBtn.Enabled = False
            End If
        ElseIf ListTable.PageIndex = ListTable.PageCount - 1 Then
            BackLBtn.Enabled = True
            NextLBtn.Enabled = False
        Else
            BackLBtn.Enabled = True
            NextLBtn.Enabled = True
        End If

    End Sub

    Protected Sub ListTable_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles ListTable.PageIndexChanging

        ListTable.PageIndex = e.NewPageIndex
        PageNumberDDL.SelectedIndex = e.NewPageIndex
        Call BackNextEnable()

    End Sub

    Protected Sub PageNumberDDL_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PageNumberDDL.SelectedIndexChanged

        ListTable.PageIndex = Convert.ToInt32(PageNumberDDL.SelectedItem.Text) - 1
        PageNumberDDL.SelectedIndex = Convert.ToInt32(PageNumberDDL.SelectedItem.Text) - 1
        Call BackNextEnable()

    End Sub

    Protected Sub PageSizeDDL_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PageSizeDDL.SelectedIndexChanged

        ListTable.PageSize = Convert.ToInt32(PageSizeDDL.SelectedValue)
        PageNumberDDL.Items.Clear()
        Call GetRecordCount()
        Call SetPage()
        ListTable.PageIndex = 0
        BackLBtn.Enabled = True
        NextLBtn.Enabled = True
        ListTable.DataBind()
        Call BackNextEnable()

    End Sub

    Protected Sub BackLBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackLBtn.Click

        If ListTable.PageIndex > 0 Then

            ListTable.PageIndex -= 1
            PageNumberDDL.SelectedIndex = ListTable.PageIndex
            BackNextEnable()

        End If

    End Sub

    Protected Sub NextLBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NextLBtn.Click

        If ListTable.PageIndex < (ListTable.PageCount - 1) Then

            ListTable.PageIndex += 1
            PageNumberDDL.SelectedIndex = ListTable.PageIndex
            BackNextEnable()

        End If

    End Sub
End Class
