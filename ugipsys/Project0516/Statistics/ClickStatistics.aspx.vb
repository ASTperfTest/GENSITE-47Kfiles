Imports System.Data
Imports System.Data.SqlClient

Partial Class Statistics_ClickStatistics
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

        Dim SQL As String = "Select *, 0 as [總計] from CounterByDate Where YMD >= '" & CDate(Me.txtStartDate.Text).ToString("yyyy-MM-dd") & "' and YMD < '" & CDate(Me.txtEndDate.Text).AddDays(1).ToString("yyyy-MM-dd") & "' Order by YMD"

        'Response.Write(SQL)

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

            TotalRow("YMD") = "總計"

            TempTable.Rows.InsertAt(TotalRow, 0)
            For F As Int32 = 1 To TempTable.Columns.Count - 1
                TotalRow(F) = 0
            Next

            Dim K As Int32 = 1

            For I As Int32 = 0 To DateDiff(DateInterval.Day, CDate(Me.txtStartDate.Text), CDate(Me.txtEndDate.Text))
                Dim CurrentDate As String = CDate(Me.txtStartDate.Text).AddDays(I).ToString("yyyy-MM-dd")

                'Response.Write(K & ", " & TempTable.Rows(K)("YMD") & ", " & CurrentDate & "<br>")
                If TempTable.Rows.Count > K AndAlso TempTable.Rows(K)("YMD") = CurrentDate Then
                    For F As Int32 = 1 To TempTable.Columns.Count - 2
                        If Not IsDBNull(TempTable.Rows(K)(F)) Then
                            TempTable.Rows(0)(F) += TempTable.Rows(K)(F)
                            TempTable.Rows(0)(TempTable.Columns.Count - 1) += TempTable.Rows(K)(F)
                            TempTable.Rows(K)(TempTable.Columns.Count - 1) += TempTable.Rows(K)(F)
                        End If
                    Next
                Else
                    Dim NewRow As DataRow = TempTable.NewRow
					For Y As Int32 = 1 To TempTable.Columns.Count - 1
						NewRow(Y) = 0
					Next
                    NewRow("YMD") = CurrentDate
                    TempTable.Rows.InsertAt(NewRow, K)
                End If

                TempTable.Rows(K)(0) = "<nobr>" & CDate(TempTable.Rows(K)(0)).ToString("MM/dd") & "</nobr>"
                K += 1
            Next


            TempTable.Columns(0).ColumnName = "日期"

            Me.GridViewOfQueryResult.DataSource = TempTable

            Me.GridViewOfQueryResult.DataBind()

            '日期不要換行
            For Each GR As GridViewRow In Me.GridViewOfQueryResult.Rows
                GR.Cells(0).Text = Server.HtmlDecode(GR.Cells(0).Text)
            Next
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Me.Request.UrlReferrer Is Nothing OrElse Me.Request.UrlReferrer.ToString.ToLower.IndexOf("kmwebsys") = -1 Then
            Me.Response.End()
        End If
    End Sub
End Class
