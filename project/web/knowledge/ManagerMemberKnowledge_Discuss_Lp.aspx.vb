Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class ManagerMemberKnowledge_Discuss_Lp
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim sb As StringBuilder
        Dim MemberId As String
       

        If Request.QueryString("MemberId") IsNot Nothing Then
            MemberId = Request.QueryString("MemberId")
        Else
            MemberId = ""
        End If

        '回前頁
        If Request.QueryString("BackPageNumber") IsNot Nothing And Request.QueryString("BackPageSize") IsNot Nothing Then
            LabelBackLink.Text = "<a href=""/knowledge/KnowledgeActivityRankDetail.aspx?PageNumber=" & Request.QueryString("BackPageNumber") & "&PageSize=" & Request.QueryString("BackPageSize") & """ title=""回前頁"">回前頁</a>"
        Else
            LabelBackLink.Text = "<a href=""/knowledge/KnowledgeActivityRankDetail.aspx""  title=""回前頁"">回前頁</a>"
        End If

        
        sb = Nothing
        '-----------------------------------------------------------------------------

        Dim PageNumber As Integer
        Dim PageSize As Integer

        If Not Page.IsPostBack Then
            Try
                If Request.QueryString("PageSize") = "" Then
                    PageSize = 10
                Else
                    PageSize = Integer.Parse(Request.QueryString("PageSize"))
                End If
                If Request.QueryString("PageNumber") = "" Then
                    PageNumber = 1
                Else
                    PageNumber = Integer.Parse(Request.QueryString("PageNumber"))
                End If
            Catch ex As Exception
                If Request.QueryString("debug") = "true" Then
                    Response.Write(ex.ToString)
                    Response.End()
                End If
                Response.Redirect("/knowledge/ManagerMemberKnowledge_Discuss_Lp.aspx")
                Response.End()
            End Try
        Else
            PageSize = PageSizeDDL.SelectedValue
            PageNumber = PageNumberDDL.SelectedValue
            
        End If
        '-----------------------------------------------------------------------------

        '-----------------------------------------------------------------------------

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString
        Dim sqlString As String = ""
        Dim myConnection As SqlConnection
        Dim myCommand As SqlCommand
        Dim myReader As SqlDataReader
        Dim myTable As DataTable
        Dim myDataRow As DataRow

        Dim Total As Integer = 0
        Dim PageCount As Integer = 0
        Dim Position As Integer = 1

        myConnection = New SqlConnection(ConnString)
        Try
            sqlString = "SELECT DISTINCT KnowledgeForum_1.Status,KnowledgeForum_1.ParentIcuitem, CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.dEditDate, CuDTGeneric.xPostDate, CodeMain.mValue, KnowledgeForum.DiscussCount, "
            sqlString &= "KnowledgeForum.CommandCount, KnowledgeForum.BrowseCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, CuDTGeneric.topCat, "
            sqlString &= "KnowledgeForum.HavePros, CuDTGeneric.fCTUPublic FROM CuDTGeneric AS CuDTGeneric_1 INNER JOIN "
            sqlString &= "CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN KnowledgeForum AS KnowledgeForum_1 "
            sqlString &= "ON CuDTGeneric.iCUItem = KnowledgeForum_1.ParentIcuitem ON CuDTGeneric_1.iCUItem = KnowledgeForum_1.gicuitem INNER JOIN "
            sqlString &= "CodeMain ON CuDTGeneric.topCat = CodeMain.mCode "
            sqlString &= "INNER JOIN dbo.KnowledgeActivity ON CuDTGeneric_1.iCUItem = dbo.KnowledgeActivity.CUItemId "
            sqlString &= "WHERE (CuDTGeneric_1.iCTUnit = @ictunit) AND (CuDTGeneric_1.iEditor = @account) "
            sqlString &= "AND (CodeMain.codeMetaID = 'KnowledgeType') AND (CuDTGeneric_1.siteId = @siteid) AND (KnowledgeForum_1.Status = 'N') "
			' AND (KnowledgeForum.Status = 'N') AND ( CuDTGeneric_1.fCTUPublic = 'Y' ) "
			sqlString &= "ORDER BY CuDTGeneric.xPostDate DESC"
			'response.write(sqlString)
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)

            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeDiscussCtUnitId"))
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
            myCommand.Parameters.AddWithValue("@account", MemberId)

            myReader = myCommand.ExecuteReader()

            myTable = New DataTable()
            myTable.Load(myReader)

            Total = myTable.Rows().Count()

            If Total <> 0 Then

                PageCount = Int(Total / PageSize + 0.999)

                If PageCount < PageNumber Then
                    PageNumber = PageCount
                End If

                PageNumberText.Text = PageNumber.ToString()
                TotalPageText.Text = PageCount.ToString()
                TotalRecordText.Text = Total.ToString()
                If PageSize = 10 Then
                    PageSizeDDL.SelectedIndex = 0
                ElseIf PageSize = 20 Then
                    PageSizeDDL.SelectedIndex = 1
                ElseIf PageSize = 30 Then
                    PageSizeDDL.SelectedIndex = 2
                ElseIf PageSize = 50 Then
                    PageSizeDDL.SelectedIndex = 3
                End If
                Dim item As ListItem
                Dim j As Integer = 0
                PageNumberDDL.Items.Clear()
                For j = 0 To PageCount - 1
                    item = New ListItem
                    item.Value = j + 1
                    item.Text = j + 1
                    If PageNumber = (j + 1) Then
                        item.Selected = True
                    End If
                    PageNumberDDL.Items.Insert(j, item)
                    item = Nothing
                Next

                Position = PageSize * (PageNumber - 1)

                myDataRow = myTable.Rows.Item(Position)

                If PageNumber > 1 Then
                    PreviousText.Visible = True
                    PreviousImg.Visible = True
                    PreviousLink.NavigateUrl = "/knowledge/ManagerMemberKnowledge_Discuss_Lp.aspx?MemberId=" & MemberId & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                Else
                    PreviousText.Visible = False
                    PreviousImg.Visible = False
                End If
                If PageNumberDDL.SelectedValue < PageCount Then
                    NextText.Visible = True
                    NextImg.Visible = True
                    NextLink.NavigateUrl = "/knowledge/ManagerMemberKnowledge_Discuss_Lp.aspx?MemberId=" & MemberId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                Else
                    NextText.Visible = False
                    NextImg.Visible = False
                End If

                Dim i As Integer = 0
                Dim link As String = ""
				Dim kmwebsysSite As String = WebConfigurationManager.AppSettings("kmwebsysSite") 

                sb = New StringBuilder
                sb.Append("<div class=""lp"">")
                sb.Append("<table width=""100%"" border=""1"" cellspacing=""1"" cellpadding=""1"" summary=""問題列表資料表格"">")
                sb.Append("<tr><th scope=""col"">狀態</th><th scope=""col"">問題標題</th><th scope=""col"">討論得分</th><th scope=""col"">知識分類</th><th scope=""col"">建立日期</th><th scope=""col"">更新日期</th>")
                sb.Append("<th scope=""col"">討論</th><th scope=""col"">瀏覽</th><th scope=""col"">專家</th></tr>")

                Dim discussGrade As String
				Dim totalDiscuss As String
                For i = 0 To PageSize - 1

                    sb.Append("<tr>")
                    link = ""
                    If myDataRow("fCTUPublic") = "Y" Then
						If myDataRow("Status") = "D" Then
							sb.Append("<td>此討論已刪除</td>")
						Else
							sb.Append("<td>&nbsp;</td>")
						End If
                    ElseIf myDataRow("fCTUPublic") = "N" Then
						If myDataRow("Status") = "D" Then
							sb.Append("<td>此問題已刪除</td>")
						Else
							sb.Append("<td>此問題不公開</td>")
						End If
                    End If
					
					link = "<a href=""http://" & kmwebsysSite & "/KnowledgeForum/KnowledgeDiscussList.asp?activity=Y&nowPage=1&pagesize=15&questionId=" & myDataRow("iCUItem") & """ target=""_blank"">"
                    link &= myDataRow("sTitle") + "</a>"

                    sb.Append("<td>" & link & "</td>")

                    '計算討論分數 
                    discussGrade = "0"
                    Try
                        sqlString = "SELECT ISNULL(SUM(ka.Grade),0) FROM dbo.KnowledgeActivity KA "
                        sqlString &= "INNER JOIN dbo.KnowledgeForum K ON ka.CUItemId = K.gicuitem "
                        sqlString &= "WHERE K.ParentIcuitem=@ParentIcuitem AND KA.MemberId=@MemberId AND KA.State=1 "
                        sqlString &= "AND KA.type BETWEEN 3 AND 4 AND K.Status='N' "

                        'myConnection.Open()
                        myCommand = New SqlCommand(sqlString, myConnection)
                        myCommand.Parameters.AddWithValue("@ParentIcuitem", myDataRow("ParentIcuitem"))
                        myCommand.Parameters.AddWithValue("@MemberId", MemberId)

                        discussGrade = CType(myCommand.ExecuteScalar, String)


                        myCommand.Dispose()

                        If Integer.Parse(discussGrade) > 4 Then
                            discussGrade = "4"
                        End If

                    Catch ex As Exception
                        If Request.QueryString("debug") = "true" Then
                            Response.Write(ex.ToString)
                            Response.End()
                        End If
                        Response.Redirect("/knowledge/ManagerMemberKnowledge_Discuss_Lp.aspx")
                        Response.End()
                    Finally
                        If myConnection.State = ConnectionState.Open Then
                            'myConnection.Close()
                        End If
                    End Try
                    '計算討論分數 end

                    sb.Append("<td>" & discussGrade & "</td>")
                    sb.Append("<td>" & myDataRow("mValue") & "</td>")
                    sb.Append("<td>[" & Date.Parse(myDataRow("xPostDate")).ToShortDateString & "]</td>")
                    sb.Append("<td>[" & Date.Parse(myDataRow("dEditDate")).ToShortDateString & "]</td>")
                    
					'計算討論數
					totalDiscuss = "0"
					Dim knowledgeUtility As New KnowledgeUtility(myDataRow("iCUItem"))
					totalDiscuss = knowledgeUtility.GetDiscussCount()
					sb.Append("<td>" & totalDiscuss & "</td>")
					'sb.Append("<td>" & myDataRow("DiscussCount") & "</td>")

                    sb.Append("<td>" & myDataRow("BrowseCount") & "</td>")
                    If myDataRow("HavePros") = "Y" Then
                        sb.Append("<td>V</td>")
                    Else
                        sb.Append("<td>&nbsp;</td>")
                    End If

                    

                    sb.Append("</tr>")

                    Position += 1
                    If myTable.Rows.Count <= Position Then
                        Exit For
                    End If
                    myDataRow = myTable.Rows(Position)
                Next

                sb.Append("</table></div>")

                TableText.Text = sb.ToString()
            Else
			
				dim errMsg as string  = "<script>alert('查無資料。');history.go(-1);</script>"
                page.ClientScript.RegisterStartupScript(Me.GetType, "error", errMsg)
                'TableText.Text = "查無資料!!"
                'TableText.ForeColor = System.Drawing.Color.Red
				
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
