Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_LeftMenu
    Inherits System.Web.UI.UserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection = New SqlConnection(ConnString)
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader

        '---KnowledgeSummaryText---    
        Dim QuestionCount As Integer = 0
        Dim DiscussCount As Integer = 0
        Dim ExpertAddCount As Integer = 0
        Dim MemberCount As Integer = 0
        Dim KnowledgeTypeA As Integer = 0
        Dim KnowledgeTypeB As Integer = 0
        Dim KnowledgeTypeC As Integer = 0
        Dim KnowledgeTypeD As Integer = 0
        Dim KnowledgeTypeE As Integer = 0
        'Dim KnowledgeTypeF As Integer = 0
        Dim KnowledgeTypeTotal As Integer = 0

        Try
            myConnection.Open()

            KnowledgeSummaryText.Text = "<h2>知識統計</h2>"

            sqlString = "SELECT * FROM KnowledgeCount WHERE id = 1"

            myCommand = New SqlCommand(sqlString, myConnection)

            myReader = myCommand.ExecuteReader

            If myReader.Read Then

                QuestionCount = myReader("QuestionCount")
                DiscussCount = myReader("DiscussCount")
                ExpertAddCount = myReader("ExpertAddCount")
                MemberCount = myReader("MemberCount")
                KnowledgeTypeA = myReader("KnowledgeTypeA")
                KnowledgeTypeB = myReader("KnowledgeTypeB")
                KnowledgeTypeC = myReader("KnowledgeTypeC")
                KnowledgeTypeD = myReader("KnowledgeTypeD")
                KnowledgeTypeE = myReader("KnowledgeTypeE")
                'KnowledgeTypeF = myReader("KnowledgeTypeF")

                KnowledgeTypeTotal = KnowledgeTypeA + KnowledgeTypeB + KnowledgeTypeC + KnowledgeTypeD + KnowledgeTypeE '+ KnowledgeTypeF

            End If
            
            KnowledgeSummaryText.Text &= "<p>知識問題數：" & QuestionCount.ToString() & "</p>"

            KnowledgeSummaryText.Text &= "<p>平均討論數：" & Math.Round((DiscussCount / QuestionCount), 2).ToString() & "</p>"

            KnowledgeSummaryText.Text &= "<p>專家補充數：" & ExpertAddCount.ToString() & "</p>"

            KnowledgeSummaryText.Text &= "<p>會員總數：" & MemberCount.ToString() & "</p>"

            '---knowledge type link---
            KnowledgeTypeLink.Text = "<li><a href=""knowledge_lp.aspx?ArticleType=&CategoryId=A"" title=""農"" alt=""農"">農 (" & KnowledgeTypeA & ")</a></li>"

            KnowledgeTypeLink.Text &= "<li><a href=""knowledge_lp.aspx?ArticleType=&CategoryId=B"" title=""林"" alt=""林"">林 (" & KnowledgeTypeB & ")</a></li>"

            KnowledgeTypeLink.Text &= "<li><a href=""knowledge_lp.aspx?ArticleType=&CategoryId=C"" title=""漁"" alt=""漁"">漁 (" & KnowledgeTypeC & ")</a></li>"

            KnowledgeTypeLink.Text &= "<li><a href=""knowledge_lp.aspx?ArticleType=&CategoryId=D"" title=""牧"" alt=""牧"">牧 (" & KnowledgeTypeD & ")</a></li>"

            KnowledgeTypeLink.Text &= "<li><a href=""knowledge_lp.aspx?ArticleType=&CategoryId=E"" title=""其它"" alt=""其它"">其它 (" & KnowledgeTypeE & ")</a></li>"

            'KnowledgeTypeLink.Text &= "<li><a href=""knowledge_lp.aspx?ArticleType=&CategoryId=F"" title=""產銷經營管理系統"" alt=""產銷經營管理系統"">產銷經營管理系統 (" & KnowledgeTypeF & ")</a></li>"

            KnowledgeTypeLink.Text &= "<li><a href=""knowledge_lp.aspx?ArticleType=&CategoryId="" title=""全部"" alt=""全部"">全部 (" & KnowledgeTypeTotal & ")</a></li>"

            '---adrotatelink---
            sqlString = "SELECT TOP 5 CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeNode.CtNodeID, CuDTGeneric.showType, CuDTGeneric.xURL, CuDTGeneric.xNewWindow, "
            sqlString &= "CuDTGeneric.fileDownLoad, CuDTGeneric.xImgFile FROM CuDTGeneric INNER JOIN CatTreeNode ON CuDTGeneric.iCTUnit = CatTreeNode.CtUnitID "
            sqlString &= "WHERE (CuDTGeneric.iCTUnit = @CtUnitId) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CatTreeNode.inUse = 'Y') "
            sqlString &= "ORDER BY NEWID()"
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@CtUnitId", WebConfigurationManager.AppSettings("AdRotateCtUnitId"))
            myReader = myCommand.ExecuteReader
            AdRotateLink.Text = ""
            While myReader.Read()
                If myReader("showType") = "1" Then
                    AdRotateLink.Text &= "<li><a href=""/ct.asp?xItem=" & myReader("iCUItem") & "&ctNode=" & myReader("CtNodeID") & "&mp=1"" title=""" & myReader("sTitle") & """ alt=""" & myReader("sTitle") & """ "
                ElseIf myReader("showType") = "2" Then
                    AdRotateLink.Text &= "<li><a href=""" & myReader("xURL") & """ title=""" & myReader("sTitle") & """ alt=""" & myReader("sTitle") & """ "
                ElseIf myReader("showType") = "3" Then
                    AdRotateLink.Text &= "<li><a href=""/public/Data/" & myReader("fileDownLoad") & """ title=""" & myReader("sTitle") & """ alt=""" & myReader("sTitle") & """ "
                End If
                If myReader("xNewWindow") = "Y" Then
                    AdRotateLink.Text &= "target=""_blank"" "
                End If
                AdRotateLink.Text &= "><img src=""/public/Data/" & myReader("xImgFile") & """/></a></li>"
            End While

            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

        Catch ex As Exception
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Sub

End Class
