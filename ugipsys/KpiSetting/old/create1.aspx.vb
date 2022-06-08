Imports System.Data
Imports System.Data.SqlClient
Partial Class create1
    Inherits System.Web.UI.Page
    

    Dim cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
    Dim Rank0_2 As String








    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

 '


    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Page.IsValid Then
            Dim Rank0_2 As String
            Dim Rank0_4 As String = Request.QueryString("id")
            cn.Open()

            Dim sql As String = "select * from kpi_set_score where Rank0_4=@Rank0_4 "

            Dim da As SqlDataAdapter = New SqlDataAdapter(sql, cn)
            da.SelectCommand.Parameters.AddWithValue("@Rank0_4", Rank0_4)


            Dim ds As New DataSet




            da.Fill(ds, "kpi_set_score")
            Dim a As Integer = ds.Tables("kpi_set_score").Rows.Count
            'Response.Write(a)
            Dim i As Integer
            Dim sStr As New StringBuilder
            Dim subs As String
            Dim ss(ds.Tables("kpi_set_score").Rows.Count - 1) As String
            Dim s1(ds.Tables("kpi_set_score").Rows.Count - 1) As String
            For i = 0 To ds.Tables("kpi_set_score").Rows.Count - 1
                subs = ds.Tables("kpi_set_score").Rows(i).Item("Rank0_2")

                ss(i) = subs
            Next


            For i = 0 To ss.Length - 1

                Dim s2 As String
                Dim s3 As String
                Dim s4 As Integer
                s2 = ss(i)
                s3 = s2.Substring(5, 1)
                s4 = Integer.Parse(s3)
                s1(i) = s4

            Next
            Array.Sort(s1)
            Dim Rank As Integer = s1(s1.Length - 1) + 1
            Dim subss As String = Rank.ToString
            Rank0_2 = Rank0_4 + subss





            Dim Rank0_0 As String = Me.TextBox1.Text
            Dim Rank0_1 As String = Me.TextBox2.Text

            ' Response.Write(Rank0_2)




            Dim sql1 As String
            sql1 = "insert into kpi_set_score (Rank0,Rank0_0,Rank0_1,Rank0_2,Rank0_3,Rank0_4) values ('st_2',@Rank0_0,@Rank0_1,'st_227','瀏覽行為特殊給點','st_22')"
            'Response.Write(sql1)

            Try
                Dim command As New SqlCommand(sql1, cn)
                'Response.Write(Rank0_2)
                command.Parameters.AddWithValue("@Rank0_0", Rank0_0)
                command.Parameters.AddWithValue("@Rank0_1", Rank0_1)
                command.Parameters.AddWithValue("@Rank0_2", Rank0_2)
                command.ExecuteNonQuery()
                Server.Transfer("kpi_set_score.aspx?id=" & Rank0_4)


                cn.Close()

            Catch ex As Exception
                'Response.Write(ex.Message)



            End Try



        End If

           

     







    End Sub


End Class
