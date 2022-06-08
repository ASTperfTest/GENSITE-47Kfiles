Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
Partial Class create
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Rank0 As String = Me.TextBox1.Text
        Dim Rank1 As String = Me.TextBox2.Text
        Dim Rank2 As String = Me.Select1.Value

        Dim cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        If Page.IsValid Then


            cn.Open()

            Dim sql As String = "select * from kpi_set_ind where Rank0=@Rank0"
            Try
                Dim command As New SqlCommand(sql, cn)
                
                Dim da As SqlDataAdapter = New SqlDataAdapter(sql, cn)
                da.SelectCommand.Parameters.AddWithValue("@Rank0", Rank0)
                Dim ds As New DataSet
                da.Fill(ds, "kpi_set_ind")

                If ds.Tables("kpi_set_ind").Rows.Count > 0 Then
                    Page.ClientScript.RegisterStartupScript(Me.GetType(), "", "alert('ID­«ÂÐ');", True)
                Else

                    Try


                        Dim sql1 As String
                        sql1 = "insert into kpi_set_ind (Rank0,Rank1,Rank2) values (@Rank0,@Rank1,@Rank2)"


                        Try
                            Dim command1 As New SqlCommand(sql1, cn)
                            command.Parameters.AddWithValue("@Rank0", Rank0)
                            command.Parameters.AddWithValue("@Rank1", Rank1)
                            command.Parameters.AddWithValue("@Rank2", Rank2)
                            command1.ExecuteNonQuery()

                            Server.Transfer("kpi_set_index.aspx")


                        Catch ex As Exception
                            Response.Write(ex.Message)
                     


                        End Try

                    Catch ex As Exception
                        Response.Write(ex.Message)
                    End Try
                End If

            Catch ex As Exception
                Response.Write(ex.Message)

            Finally
                cn.Close()
            End Try







        End If


    End Sub






End Class
