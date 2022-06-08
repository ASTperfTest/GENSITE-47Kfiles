Imports System.Data
Imports System.Data.SqlClient
Partial Class kpi_set_score
    Inherits System.Web.UI.Page

    
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim cn As New SqlConnection(ConfigurationManager.ConnectionStrings("ConnString").ConnectionString)
        Dim Rank0_4 As String = Request.QueryString("id")
        If Rank0_4 = "st_22" Then


            HyperLink1.Visible = True
            HyperLink1.NavigateUrl = "create1.aspx?id=" & Rank0_4

        End If





        If Not Page.IsPostBack Then
            cn.Open()

            'Response.Write(Rank0_4)
            Try
                Dim sql As String
                sql = "select Rank0_4,Rank0_2,Rank0_0 as '¦WºÙ',Rank0_1 as '¤À¼Æ' from kpi_set_score where Rank0_4=@Rank0_4"


                Dim da As SqlDataAdapter = New SqlDataAdapter(sql, cn)
                da.SelectCommand.Parameters.AddWithValue("@Rank0_4", Rank0_4)
                Dim ds As New DataSet




                da.Fill(ds, "kpi_set_score")

                GridView1.DataSource = ds
                GridView1.DataBind()



            Catch ex As Exception
                Response.Write(ex.Message)
            Finally
                cn.Close()

            End Try
        End If
        If Not Page.IsPostBack Then
            cn.Open()
            Dim sql As String
            sql = "select Rank0_3 from kpi_set_score where Rank0_4=@Rank0_4"
            Try
                Dim da As SqlDataAdapter = New SqlDataAdapter(sql, cn)
                da.SelectCommand.Parameters.AddWithValue("@Rank0_4", Rank0_4)
                Dim ds As DataSet = New DataSet()
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)
                da.Fill(ds, "kpi_set_score")
                Dim tempi As Integer = dt.Rows.Count

                Dim a As String = dt.Rows(0)("Rank0_3").ToString()

                ' Response.Write(a)
                Label1.Text = a

            Finally
                cn.Close()

            End Try



        End If





    End Sub


    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub
End Class
