Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_trace_add
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ArticleId As String
        Dim MemberId As String
        Dim Script As String = "<script>alert('{0}!!');window.location.href='" & Request.UrlReferrer.ToString() &"';</script>"

        '-----------------------------------------------------------------------------
        If Session("memID") = "" Or Session("memID") Is Nothing Then
            MemberId = ""
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script.Replace("{0}", "連線逾時或尚未登入，請登入會員"))
            Exit Sub
        Else
            MemberId = Session("memID").ToString()
        End If
        '-----------------------------------------------------------------------------

        If Request.QueryString("ArticleId") = "" Then
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script.Replace("{0}", "加入追蹤的文章錯誤"))
            Exit Sub
        Else
            ArticleId = Request.QueryString("ArticleId")
        End If

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim Flag As Boolean = False
        Dim QuestionIctunit As String = WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId")
        Dim MaxIcuitem As String

        myConnection = New SqlConnection(ConnString)
        Try            
            sqlString = "SELECT iCTUnit, fCTUPublic FROM CuDTGeneric INNER JOIN KnowledgeForum ON KnowledgeForum.gicuitem = CuDTGeneric.iCUItem "
            sqlString &= "WHERE iCUItem = @icuitem AND (KnowledgeForum.Status = 'N') AND (CuDTGeneric.fCTUPublic = 'Y')"
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
            myReader = myCommand.ExecuteReader()

            If myReader.Read Then
                Flag = True
                If myReader("iCTUnit") <> QuestionIctunit Then
                    Flag = False
                End If
            Else
                Flag = False
            End If

            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            If Flag = False Then
                Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script.Replace("{0}", "無文章可加入追蹤"))
                Exit Sub
            Else
                '---檢查是否加入過這文章---
                sqlString = "SELECT iCUItem FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
                sqlString &= "WHERE KnowledgeForum.ParentIcuitem = @icuitem AND CuDTGeneric.iEditor = @ieditor "
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@icuitem", ArticleId)
                myCommand.Parameters.AddWithValue("@ieditor", MemberId)
                myReader = myCommand.ExecuteReader()

                If myReader.Read Then
                    Flag = False
                Else
                    Flag = True
                End If
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

                If Not Flag Then
                    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script.Replace("{0}", "此文章已加入追蹤"))
                    Exit Sub
                ElseIf Flag Then

                    '---加入追蹤---
                    sqlString = "INSERT INTO CuDTGeneric(iBaseDSD, iCTUnit, fCTUPublic, sTitle, iEditor, iDept, siteId) VALUES (@ibasedsd, @ictunit, 'Y', @stitle, @ieditor, '0', @siteid)"
                    myCommand = New SqlCommand(sqlString, myConnection)
                    myCommand.Parameters.Add(New SqlParameter("@ibasedsd", WebConfigurationManager.AppSettings("KnowledgeBaseDSD")))
                    myCommand.Parameters.Add(New SqlParameter("@ictunit", WebConfigurationManager.AppSettings("KnowledgeTraceCtUnitId")))
                    myCommand.Parameters.Add(New SqlParameter("@stitle", "加入追蹤-" & ArticleId))
                    myCommand.Parameters.Add(New SqlParameter("@ieditor", MemberId))
                    myCommand.Parameters.Add(New SqlParameter("@siteid", WebConfigurationManager.AppSettings("SiteId")))
                    myCommand.ExecuteNonQuery()
                    myCommand.Dispose()

                    sqlString = "SELECT @@identity AS MaxIcuitem"
                    Dim MaxCommand As SqlDataAdapter = New SqlDataAdapter(sqlString, myConnection)
                    Dim MaxDataSet As DataSet = New DataSet()
                    MaxCommand.Fill(MaxDataSet, "IndexRecommend")
                    Dim myTable As DataTable = MaxDataSet.Tables("IndexRecommend")
                    MaxIcuitem = myTable.Rows(0).Item(0)

                    sqlString = "INSERT INTO KnowledgeForum(gicuitem, ParentIcuitem) VALUES (@gicuitem, @parenticuitem)"
                    myCommand = New SqlCommand(sqlString, myConnection)
                    myCommand.Parameters.Add(New SqlParameter("@gicuitem", MaxIcuitem))
                    myCommand.Parameters.Add(New SqlParameter("@parenticuitem", ArticleId))
                    myCommand.ExecuteNonQuery()
                    myCommand.Dispose()

                    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "TraceSuccess", Script.Replace("{0}", "文章加入追蹤成功"))
                End If
            End If

        Catch ex As Exception

            If Request("debug") = "true" Then
                Response.Write(ex.ToString())
                Response.End()
            End If
            Response.Redirect("/knowledge/knowledge.aspx")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Sub

End Class
