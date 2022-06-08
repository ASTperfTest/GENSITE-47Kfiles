Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Data.SqlClient
Imports System.Text

Partial Class QuestionResponse
    Inherits System.Web.UI.Page
	dim todo as string ="尚未處理"
    Dim done As String = "處理完畢"
    Dim del As String = "刪除"
    Dim objPage As PagedDataSource

    Protected Sub GridView1_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs) Handles GridView3.PageIndexChanging
        
    End Sub

    Protected Sub GridView1_PageIndexChanging_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs) Handles GridView3.RowEditing
        Me.Panel1.Visible = True

        Dim dv As GridView = CType(sender, GridView)
        For i As Integer = 0 To dv.Columns.Count - 1
            Dim fieldString As String = CType(sender, GridView).Rows(e.NewEditIndex).Cells(i ).Text
            Select Case dv.Columns.Item(i).HeaderText
                Case "問題單號"
                    Me.txtSEQ.Text = HttpUtility.HtmlDecode(fieldString)
                Case "處理敘述"
                    Me.textboxRESPONSE.Text = HttpUtility.HtmlDecode(fieldString)
                Case "處理狀態"
                    'jira 問題單回應 後台管理 http://gssjira.gss.com.tw/browse/COAKM-49

                    Select Case fieldString
                        Case todo
                            Me.drpSTATUS.SelectedValue = "0"
                        Case done
                            Me.drpSTATUS.SelectedValue = "1"
                        Case del
                            Me.drpSTATUS.SelectedValue = "9"
                    End Select
                    'Case Else
                    'Response.Write(fieldString + "<br/>")
            End Select
        Next


        'Me.txtSEQ.Text = HttpUtility.HtmlDecode(CType(sender, GridView).Rows(e.NewEditIndex).Cells(1).Text).Trim
        'Me.textboxRESPONSE.Text = HttpUtility.HtmlDecode(CType(sender, GridView).Rows(e.NewEditIndex).Cells(7).Text).Trim
        'Me.drpSTATUS.SelectedItem.Text = CType(sender, GridView).Rows(e.NewEditIndex).Cells(8).Text
    End Sub

    Protected Sub buttonSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonSave.Click
        Dim connection2 As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
        Try
            Using (connection2)
                connection2.Open()
                Dim sql As StringBuilder = New StringBuilder
                sql.Append(" UPDATE KNOWLEDGE_REPORT ")
                sql.Append(" SET RESPONSE=@RESPONSE, STATUS=@STATUS, LAST_MODIFIER=@LAST_MODIFIER, LAST_MODIFY_DATETIME=GetDate()")
                sql.Append(" WHERE SEQ = @SEQ")
                ' Dim Cmmd As SqlCommand = New SqlCommand(sql.ToString, connection2)
                Dim Cmmd As SqlCommand = New SqlCommand(sql.ToString, connection2)
				dim memID as String=""
				If(Session("Name") is Nothing) Then
					memID="管理者"
				Else
					memID=Session("Name")
				end if
                Cmmd.Parameters.AddWithValue("@LAST_MODIFIER", memID)
                Cmmd.Parameters.AddWithValue("@RESPONSE", Me.textboxRESPONSE.Text.Trim)
                Cmmd.Parameters.AddWithValue("@STATUS", Me.drpSTATUS.SelectedValue)
                Cmmd.Parameters.AddWithValue("@SEQ", Me.txtSEQ.Text)
                Cmmd.ExecuteNonQuery()

                SelectData()
                MyBindData()
                Page.ClientScript.RegisterStartupScript(Me.Page.GetType, "updateSuccess", "alert('修改成功')", True)
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Protected Sub buttonClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles buttonClose.Click
		Me.Panel1.Visible = False

    End Sub
    Private dt As DataTable = New DataTable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        SelectData()

        If (ViewState("page") = Nothing) Then
            ViewState("page") = 0
        End If

        InitClientPageItem()
        If Page.IsPostBack = False Then
            MyBindData()
        End If
    End Sub
    Private Sub SelectData()
        Dim gridDt As DataTable = New DataTable
        Dim connection As SqlConnection = New SqlConnection(ConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString)
        Using (connection)

            Dim Cmmd As SqlCommand = connection.CreateCommand()
            Cmmd.CommandText = "SELECT * FROM [KNOWLEDGE_REPORT]"
            Cmmd.CommandText += vbCrLf + " where (@type  ='' or @type  =[type])"
            Cmmd.CommandText += vbCrLf + " and   (@status='' or @status=[status])"
            Cmmd.CommandText += vbCrLf + " and   (@description='' or [description] like '%' + @description + '%' )"
            Cmmd.CommandText += vbCrLf + " and   (@sender=''      or [creator] in"
            Cmmd.CommandText += vbCrLf + "              ("
            Cmmd.CommandText += vbCrLf + "                  select account from Member "
            Cmmd.CommandText += vbCrLf + "                  where account   like '%' + @sender + '%' "
            Cmmd.CommandText += vbCrLf + "                  or    realname  like '%' + @sender + '%' "
            Cmmd.CommandText += vbCrLf + "                  or    nickname  like '%' + @sender + '%' "
            Cmmd.CommandText += vbCrLf + "              )"
            Cmmd.CommandText += vbCrLf + "       )"
            Cmmd.CommandText += vbCrLf + " order by SEQ desc"

            Cmmd.Parameters.AddWithValue("@type", searchSource.SelectedValue)
            Cmmd.Parameters.AddWithValue("@status", searchSTATUS.SelectedValue)
            Cmmd.Parameters.AddWithValue("@description", Me.searchDescription.Text.Trim())
            Cmmd.Parameters.AddWithValue("@sender", Me.searchSender.Text.Trim())
            
            Dim Da As SqlDataAdapter = New SqlDataAdapter(Cmmd)
            Da.Fill(gridDt)

        End Using

        If objPage Is Nothing Then
            objPage = New PagedDataSource()
        End If

        objPage.DataSource = gridDt.DefaultView
        objPage.AllowPaging = True
        objPage.PageSize = PageDDSize.SelectedValue

        If ViewState("page") >= objPage.PageCount Then
            ViewState("page") = objPage.PageCount - 1
        End If
        If ViewState("page") < 0 Then
            ViewState("page") = 0
        End If
        objPage.CurrentPageIndex = ViewState("page")
    End Sub
    Private Sub InitClientPageItem()
        If PageDDList.Items.Count <> objPage.PageCount Then
            PageDDList.Items.Clear()
        Else
            Return
        End If

        For i As Integer = 0 To objPage.PageCount - 1
            PageDDList.Items.Add(New ListItem((i + 1).ToString(), i.ToString()))
        Next
    End Sub
    Private Sub MyBindData()
        Me.GridView3.DataSource = objPage
        Me.GridView3.DataBind()

        ViewState("page") = objPage.CurrentPageIndex
        PageDDList.SelectedValue = objPage.CurrentPageIndex.ToString()
        PageInfo.Text = String.Format("共<em>{0}</em>筆資料，目前在第<em>{1}/{2}</em>頁每頁顯示", objPage.DataSourceCount, objPage.CurrentPageIndex + 1, objPage.PageCount)
        Panel1.Visible = False
    End Sub
    Protected Sub GridView3_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView3.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim dv As GridView = CType(sender, GridView)
            For i As Integer = 0 To dv.Columns.Count - 1
                Select Case dv.Columns.Item(i).HeaderText
                    Case "問題來源"
                        Select Case e.Row.Cells(i).Text
                            Case "1"
                                e.Row.Cells(i).Text = "問題反應-<br/>知識家"
                            Case "2"
                                e.Row.Cells(i).Text = "討論檢舉-<br/>知識家"
                            Case "3"
                                e.Row.Cells(i).Text = "問題反應-<br/>知識庫"
                            Case "4"
                                e.Row.Cells(i).Text = "問題反應-<br/>主題館"
                            Case "99"
                                e.Row.Cells(i).Text = "問題反應-<br/>最新消息"
                        End Select
                    Case "處理狀態"
                        Select Case e.Row.Cells(i).Text
                            Case "0"
                                e.Row.Cells(i).Text = todo
                            Case "1"
                                e.Row.Cells(i).Text = done
                            Case "9"
                                e.Row.Cells(i).Text = del
                        End Select
                        'Case Else
                        'Response.Write(dv.Columns.Item(i).HeaderText + i.ToString() + "<br/>")
                End Select
            Next

        End If

    End Sub
    Protected Sub OnClick_Back(ByVal sender As Object, ByVal e As EventArgs)

        Dim newpage As Integer = CType(ViewState("page"), Integer) = 1
        If newpage < 0 Then
            objPage.CurrentPageIndex = 0
        Else
            objPage.CurrentPageIndex = newpage
        End If
        MyBindData()
    End Sub
    Protected Sub OnClick_Next(ByVal sender As Object, ByVal e As EventArgs)

        Dim newpage As Integer = CType(ViewState("page"), Integer) + 1
        If newpage >= objPage.PageCount Then
            objPage.CurrentPageIndex = objPage.PageCount - 1
        Else
            objPage.CurrentPageIndex = newpage
        End If
        MyBindData()
    End Sub
    Protected Sub OnSetNewPagesize(ByVal sender As Object, ByVal e As EventArgs)
        objPage.PageSize = PageDDSize.SelectedValue
        MyBindData()
    End Sub
    Protected Sub OnGoNewPage(ByVal sender As Object, ByVal e As EventArgs)
        objPage.CurrentPageIndex = CType(sender, DropDownList).SelectedIndex
        MyBindData()
    End Sub

    Protected Sub searchSubmit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles searchSubmit.Click
        Me.SelectData()
        Me.MyBindData()
    End Sub
End Class
