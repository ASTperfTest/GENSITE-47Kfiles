Imports System.Data
Imports System.Data.SqlClient
Partial Class score_ed
    Inherits System.Web.UI.Page

    Dim cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim Rank0_2 As String = Request.QueryString("id")



        If Not Page.IsPostBack Then
            cn.Open()
            Dim sql As String
            sql = "select Rank0_0 from kpi_set_score where Rank0_2=@Rank0_2"
            Try

                Dim da As SqlDataAdapter = New SqlDataAdapter(sql, cn)
                da.SelectCommand.Parameters.AddWithValue("@Rank0_2", Rank0_2)
                Dim ds As DataSet = New DataSet()
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)
                da.Fill(ds, "kpi_set_score")
                Dim tempi As Integer = dt.Rows.Count

                Dim a As String = dt.Rows(0)("Rank0_0").ToString()

                ' Response.Write(a)
                Label1.Text = a
            Catch ex As Exception
            Finally
                cn.Close()

            End Try



        End If












    End Sub




    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Rank0_2 As String = Request.QueryString("id")
        Dim Rank0_1 As String = Me.TextBox2.Text
        Dim sql As String
        If Page.IsValid Then




            cn.Open()

            Try
                sql = "update dbo.kpi_set_score set Rank0_1 =@Rank0_1 where Rank0_2=@Rank0_2"
                Dim Rank0_4 As String = Request.QueryString("sid")

                Dim command As New SqlCommand(sql, cn)
                command.Parameters.AddWithValue("@Rank0_1", Rank0_1)
                command.Parameters.AddWithValue("@Rank0_2", Rank0_2)
                command.ExecuteNonQuery()
                Server.Transfer("kpi_set_score.aspx?id=" & Rank0_4)



                '




            Catch ex As Exception
                ' Response.Write(ex.Message)
            Finally
                cn.Close()
            End Try
        End If
 

    End Sub


  
End Class
