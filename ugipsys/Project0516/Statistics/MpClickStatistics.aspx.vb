Imports System.Data
Imports System.Data.SqlClient

Partial Class Statistics_MpClickStatistics
    Inherits System.Web.UI.Page

    Function GetTotlaCount() 

        '檢查日期

        Dim SQL As String = "Select 0 AS 首頁,SUM(最新消息) AS 最新消息,SUM(農業與生活)AS 農業與生活,SUM(新優質農家) AS 新優質農家, "
        SQL = SQL + " SUM(產銷專欄) AS 產銷專欄,SUM(資源推薦) AS 資源推薦,SUM(影音專區) AS 影音專區, SUM(相關網站) "
        SQL = SQL + " AS 相關網站, SUM(農業小百科) AS 農業小百科,SUM(農業知識家) AS 農業知識家 , "
        SQL = SQL + " SUM(農業知識庫) AS 農業知識庫,SUM(主題館) AS 主題館,SUM(農作物地圖) AS 農作物地圖,SUM(最新消息)+SUM(農業與生活) "
        SQL = SQL + " +SUM(新優質農家)+SUM(產銷專欄)+SUM(資源推薦)+ SUM(影音專區)+SUM(相關網站)"
        SQL = SQL + " +SUM(農業小百科)+SUM(農業知識家)+SUM(農業知識庫)+SUM(農作物地圖)+SUM(主題館) AS 總計 from CounterByDate "

        'Response.Write(SQL)

        'Response.End()

        Dim webConfigConnectionString As String = System.Web.Configuration.WebConfigurationManager.ConnectionStrings("ConnectionString").ConnectionString

        Dim TempTable As New DataTable()

        Dim DA As New SqlDataAdapter(SQL, webConfigConnectionString)
        DA.Fill(TempTable)

        SQL = " SELECT counts FROM counter where mp=1 "
        Dim myConnection As SqlConnection
        myConnection = New SqlConnection(webConfigConnectionString)
        Try
            myConnection.Open()
            Dim myCommand As SqlCommand
            myCommand = New SqlCommand(SQL, myConnection)
            Dim myReader As SqlDataReader
            myReader = myCommand.ExecuteReader()
            If myReader.Read() Then
                TempTable.Rows(0)(0) = myReader("counts")
            End If
        Catch ex As Exception
        Finally
            myConnection.Close()
        End Try

        If TempTable.Rows.Count = 0 Then
            Me.lblNodata.Text = "本區間查無資料.."
            Me.GridViewOfQueryResult.Visible = False
            Me.lblNodata.Visible = True
        Else
            Me.lblNodata.Visible = False
            Me.GridViewOfQueryResult.Visible = True

            Dim k As Integer
            k = TempTable.Columns.Count - 1
            TempTable.Rows(0)(k) = TempTable.Rows(0)(k) + TempTable.Rows(0)(0)


            Me.GridViewOfQueryResult.DataSource = TempTable

            Me.GridViewOfQueryResult.DataBind()

            '日期不要換行
            For Each GR As GridViewRow In Me.GridViewOfQueryResult.Rows
                GR.Cells(0).Text = Server.HtmlDecode(GR.Cells(0).Text)
            Next
        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Me.Request.UrlReferrer Is Nothing OrElse Me.Request.UrlReferrer.ToString.ToLower.IndexOf("kmwebsys") = -1 Then
            Me.Response.End()
        End If
		GetTotlaCount()
    End Sub
End Class
