Imports System.Data
Imports System.Data.SqlClient
Partial Class kpi_set_scored
    Inherits System.Web.UI.Page

   
    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        If Not Page.IsPostBack Then
            cn.Open()
            Dim Rank0 As String = Request.QueryString("id")

            Try

                Dim sql As String


                sql = " select distinct Rank0_4,Rank0_3 as '¦WºÙ' from dbo.kpi_set_score where Rank0= @Rank0"


                Dim da As SqlDataAdapter = New SqlDataAdapter(sql, cn)
                da.SelectCommand.Parameters.AddWithValue("@Rank0", Rank0)
                Dim ds As New DataSet




                da.Fill(ds, "kpi_set_index")
                GridView1.DataSource = ds
                GridView1.DataBind()


            Catch ex As Exception
            Finally
                cn.Close()

            End Try



        End If

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

       
    End Sub
End Class
