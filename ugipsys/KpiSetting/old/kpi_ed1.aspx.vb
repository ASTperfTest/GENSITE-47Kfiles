Imports System.Data
Imports System.Data.SqlClient
Partial Class kpi_ed1
    Inherits System.Web.UI.Page
    Dim cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim Rank0 As String = Request.QueryString("id")
        'Dim Rank0_1 As String = Me.TextBox1.Text
        'Response.Write(Rank0_2)

        If Not Page.IsPostBack Then
            cn.Open()
            Try

                Dim sql As String
                'sql = "select Rank1  from kpi_set_ind where Rank0='" + Rank0 + "'"
                sql = "select Rank1  from kpi_set_ind where Rank0 =@Rank0"
                Dim da As SqlDataAdapter = New SqlDataAdapter(sql, cn)
                da.SelectCommand.Parameters.AddWithValue("@Rank0", Rank0)

                Dim dt As DataTable = New DataTable()
                da.Fill(dt)
                

                TextBox1.Text = dt.Rows(0)("Rank1").ToString()
            Catch ex As Exception

            Finally
                cn.Close()

            End Try



        End If










    End Sub
    






    

    Protected Sub Select2_ServerChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles Select2.ServerChange

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Rank2 As String = Me.Select2.Value
        Dim Rank0 As String = Request.QueryString("id")
        Dim Rank1 As String = Me.TextBox1.Text
        'Dim cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        If Page.IsValid Then
            'dim strsql as = "insert into kpi_set_ind"&"(kpi_set_id,kpi_set_charts,kpi_set_charts_detail,kpi_set_charts_condition) valuse 
            Dim sql As String

            'sql = "update kpi_set_ind set Rank2 ='" + Rank2 + "' where Rank0 ='" + Rank0 + "'"
            sql = "update kpi_set_ind set Rank2 =@Rank2,Rank1=@Rank1  where Rank0 =@Rank0"
            cn.Open()
            ' Response.Write(sql)
            Try
                Dim command As New SqlCommand(sql, cn)
                command.Parameters.AddWithValue("@Rank2", Rank2)
                command.Parameters.AddWithValue("@Rank0", Rank0)
                command.Parameters.AddWithValue("@Rank1", Rank1)
                command.ExecuteNonQuery()

                Server.Transfer("kpi_set_index.aspx")



                'Dim reader As SqlDataReader = command.ExecuteReader()
                'Response.Write("Successful")
            Catch ex As Exception
                ' Response.Write(ex.Message)
            Finally
                cn.Close()
            End Try

        End If
    End Sub

    
End Class
