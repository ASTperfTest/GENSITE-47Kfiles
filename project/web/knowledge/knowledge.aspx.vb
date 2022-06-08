Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim sb As StringBuilder
        Dim i As Integer = 0

        Using myConnection As SqlConnection = New SqlConnection(ConnString)
            Try
                myConnection.Open()

                '---主題推薦---
                sb = New StringBuilder
                sqlString = "SELECT TOP (5) iCUItem, topCat, xPostDate, sTitle, xBody, xImgFile, CodeMain.mValue FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode "
                sqlString &= "INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
                sqlString &= "WHERE (iCTUnit = @ictunit) AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.xImportant <> 0) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) "
                sqlString &= "AND (KnowledgeForum.Status = 'N') ORDER BY xImportant Desc, xPostDate DESC"
                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
                myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
                myReader = myCommand.ExecuteReader

                sb.Append("<div class=""sublayout""><h3>知識主題推薦討論</h3>")
                sb.Append("<div class=""rss""><a href=""/knowledge/knowledge_rss.aspx?type=A"">RSS訂閱</a></div>")

                While myReader.Read()

                    If i = 0 Then
                        If myReader("xImgFile") IsNot DBNull.Value Then
                            sb.Append("<img src=""/public/Data/" & myReader("xImgFile") & """>")
                        End If
                        sb.Append("<h4>" & myReader("sTitle") & "</h4><p>")
                        If myReader("xBody").ToString().Length > 100 Then
                            sb.Append(myReader("xBody").ToString().Substring(0, 100) & "...")
                        Else
                            sb.Append(myReader("xBody"))
                        End If
                        sb.Append("<a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("iCUItem").ToString() & "&ArticleType=A&CategoryId=" & myReader("topCat") & """>詳全文</a></p><ul class=""list"">")
                        i += 1
                    Else
                        sb.Append("<li><a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("iCUItem") & "&ArticleType=A&CategoryId=" & myReader("topCat") & """>")
                        sb.Append(myReader("sTitle") & "</a><span>" & myReader("mValue") & "</span>[" & Date.Parse(myReader("xPostDate")).ToShortDateString & "]</li>")
                    End If

                End While
                sb.Append("</ul>")
                sb.Append("<div class=""more""><a href=""/knowledge/knowledge_lp.aspx?ArticleType=F&CategoryId="">more...</a></div></div>")
                TopicRecommandText.Text = sb.ToString
                sb = Nothing
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

                '---最新發問---
                sb = New StringBuilder
                sqlString = "SELECT TOP (5) iCUItem, topCat, xPostDate, sTitle, CodeMain.mValue FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode "
                sqlString &= "INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
                sqlString &= "WHERE (iCTUnit = @ictunit) AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) "
                sqlString &= "AND (KnowledgeForum.Status = 'N') ORDER BY xPostDate DESC"

                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
                myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
                myReader = myCommand.ExecuteReader

                sb.Append("<div class=""hotQa""><h3>最新發問</h3>")
                sb.Append("<div class=""rss""><a href=""/knowledge/knowledge_rss.aspx?type=B"">RSS訂閱</a></div>")
                sb.Append("<ul class=""list"">")
                While myReader.Read()
                    sb.Append("<li><a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("iCUItem") & "&ArticleType=A&CategoryId=" & myReader("topCat") & """>")
                    sb.Append(myReader("sTitle") & "</a><span>" & myReader("mValue") & "</span>[" & Date.Parse(myReader("xPostDate")).ToShortDateString & "]</li>")
                End While
                sb.Append("</ul>")
                sb.Append("<div class=""more""><a href=""/knowledge/knowledge_lp.aspx?ArticleType=A&CategoryId="">more...</a></div></div>")
                LatestQuestionText.Text = sb.ToString
                sb = Nothing
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

                '---熱門討論---
                sb = New StringBuilder
                sqlString = "SELECT TOP (5) CuDTGeneric.iCUItem, CuDTGeneric.topCat, CuDTGeneric.xPostDate, CuDTGeneric.sTitle, CodeMain.mValue "
                sqlString &= "FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
                sqlString &= "WHERE (iCTUnit = @ictunit) AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) "
                sqlString &= "AND (KnowledgeForum.Status = 'N') ORDER BY KnowledgeForum.DiscussCount DESC"

                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
                myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
                myReader = myCommand.ExecuteReader

                sb.Append("<div class=""hotessay""><h3>熱門討論</h3>")
                sb.Append("<div class=""rss""><a href=""/knowledge/knowledge_rss.aspx?type=C"">RSS訂閱</a></div>")
                sb.Append("<ul class=""list"">")

                While myReader.Read
                    sb.Append("<li><a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("iCUItem") & "&ArticleType=A&CategoryId=" & myReader("topCat") & """>")
                    sb.Append(myReader("sTitle") & "</a><span>" & myReader("mValue") & "</span>[" & Date.Parse(myReader("xPostDate")).ToShortDateString & "]</li>")
                End While
                sb.Append("</ul>")
                sb.Append("<div class=""more""><a href=""/knowledge/knowledge_lp.aspx?ArticleType=B&CategoryId="">more...</a></div></div>")
                HotDiscussText.Text = sb.ToString
                sb = Nothing
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

                sb = New StringBuilder
                sqlString = "SELECT TOP (5) CuDTGeneric.iCUItem, CuDTGeneric.topCat, CuDTGeneric.xPostDate, CuDTGeneric.sTitle, CodeMain.mValue "
                sqlString &= "FROM CuDTGeneric INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
                sqlString &= "WHERE (KnowledgeForum.Status = 'N') AND (CuDTGeneric.iCTUnit = @ictunit) AND (CodeMain.codeMetaID = 'KnowledgeType') AND "
                sqlString &= "(KnowledgeForum.HavePros = 'Y') AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) ORDER BY CuDTGeneric.xPostDate DESC"

                myCommand = New SqlCommand(sqlString, myConnection)
                myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
                myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
                myReader = myCommand.ExecuteReader

                sb.Append("<div class=""hotessay2""><h3>專家補充</h3>")
                sb.Append("<div class=""rss""><a href=""/knowledge/knowledge_rss.aspx?type=D"">RSS訂閱</a></div>")
                sb.Append("<ul class=""list"">")

                While myReader.Read
                    sb.Append("<li><a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myReader("iCUItem") & "&ArticleType=A&CategoryId=" & myReader("topCat") & """>")
                    sb.Append(myReader("sTitle") & "</a><span>" & myReader("mValue") & "</span>[" & Date.Parse(myReader("xPostDate")).ToShortDateString & "]</li>")
                End While
                sb.Append("</ul>")
                sb.Append("<div class=""more""><a href=""/knowledge/knowledge_lp.aspx?ArticleType=E&CategoryId="">more...</a></div></div>")
                ProfessionalText.Text = sb.ToString
                sb = Nothing
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myCommand.Dispose()

            Catch ex As Exception
                If Request.QueryString("debug") = "true" Then
                    Response.Write(ex.ToString)
                    Response.End()
                End If
            Finally
                'If myConnection.State = ConnectionState.Open Then
                'myConnection.Close()
                'End If
            End Try

        End Using
    End Sub

End Class
