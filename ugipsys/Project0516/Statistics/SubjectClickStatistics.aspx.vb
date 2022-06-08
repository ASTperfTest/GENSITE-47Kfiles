Imports System.Data
Imports System.Data.SqlClient

Partial Class Statistics_SubjectClickStatistics
    Inherits System.Web.UI.Page

    Protected Sub btnQuery_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        '檢查日期
        If Not IsDate(Me.txtStartDate.Text) OrElse Not IsDate(Me.txtEndDate.Text) Then
            Response.Write("<script>alert('開始日期或結束日期格式有誤, 請使用正常的日期格式' )</script>")
            Return
        End If

        If CDate(Me.txtStartDate.Text) > CDate(Me.txtEndDate.Text) Then
            Response.Write("<script>alert('開始日期不可大於結束日期' )</script>")
            Return
        End If

        If CDate(Me.txtStartDate.Text).AddYears(1) < CDate(Me.txtEndDate.Text) Then
            Response.Write("<script>alert('開始與結束日期區間不可超過一年，以免影響效能。' )</script>")
            Return
        End If

        Dim SQL As String = "Select CatTreeRoot.CtRootName as [主題館名稱], A.ViewCount as [瀏覽次數] from"
        SQL &= " (Select ctRootId, sum(ViewCount) as ViewCount "
        SQL &= " from CounterForSubjectByDate"
		SQL &= " where YMD >='" & CDate(Me.txtStartDate.Text).ToString("yyyy-MM-dd") & "' and YMD <= '" & CDate(Me.txtEndDate.Text).ToString("yyyy-MM-dd") & "' and ctRootId > 1 group by ctRootId) A"
		SQL &= " Join CatTreeRoot "
        SQL &= " on CatTreeRoot.CtRootID = A.CtRootID order by A.ViewCount Desc"

        'SQL &= "Response.Write(SQL)

        'Response.End()

        Dim webConfigConnectionString As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString

        Dim TempTable As New DataTable()

        Dim DA As New SqlDataAdapter(SQL, webConfigConnectionString)
        DA.Fill(TempTable)

        If TempTable.Rows.Count = 0 Then
            Me.lblNodata.Text = "本區間查無資料.."
            Me.GridViewOfQueryResult.Visible = False
            Me.lblNodata.Visible = True
        Else
            Me.lblNodata.Visible = False
            Me.GridViewOfQueryResult.Visible = True
            Dim TotalRow As DataRow = TempTable.NewRow

            Me.GridViewOfQueryResult.DataSource = TempTable

            Me.GridViewOfQueryResult.DataBind()
        End If


    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Me.Request.UrlReferrer Is Nothing OrElse Me.Request.UrlReferrer.ToString.ToLower.IndexOf("kmwebsys") = -1 Then
            Me.Response.End()
        End If
    End Sub
End Class
