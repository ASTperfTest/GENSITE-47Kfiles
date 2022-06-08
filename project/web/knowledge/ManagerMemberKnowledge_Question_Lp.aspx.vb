Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Partial Class ManagerMemberKnowledge_Question_Lp
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



        '-----------------------------------------------------------------------------
        
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
                Response.Redirect("/knowledge/ManagerMemberKnowledge_Question_Lp.aspx")
                Response.End()
            End Try
        Else
            PageSize = PageSizeDDL.SelectedValue
            If PageNumberDDL.items.count > 0 Then
                PageNumber = PageNumberDDL.SelectedValue
            Else
                PageNumber = 1
            End If
            
        End If
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
            sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.fCTUPublic, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CodeMain.mValue, CuDTGeneric.dEditDate, KnowledgeForum.DiscussCount,KnowledgeForum.Status, "
            sqlString &= "KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, KnowledgeForum.BrowseCount, KnowledgeForum.HavePros, CuDTGeneric.topCat, KnowledgeActivity.Grade,KnowledgeActivity.CreateTime "
            sqlString &= "FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem INNER JOIN CodeMain ON CuDTGeneric.topCat = CodeMain.mCode "
            sqlString &= "INNER JOIN dbo.KnowledgeActivity ON CuDTGeneric.iCUItem = dbo.KnowledgeActivity.CUItemId "
            sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.siteId = @siteid) AND (CodeMain.codeMetaID = 'KnowledgeType') AND ( CuDTGeneric.iEditor = @account) "
            sqlString &= "ORDER BY KnowledgeActivity.CreateTime Desc ,CuDTGeneric.xPostDate Desc "
            'sqlString &= "AND (KnowledgeForum.Status = 'N') AND ( CuDTGeneric.fCTUPublic = 'Y' )"

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)

            myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
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
                    PreviousLink.NavigateUrl = "/knowledge/ManagerMemberKnowledge_Question_Lp.aspx?MemberId=" & MemberId & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                Else
                    PreviousText.Visible = False
                    PreviousImg.Visible = False
                End If
                If PageNumberDDL.SelectedValue < PageCount Then
                    NextText.Visible = True
                    NextImg.Visible = True
                    NextLink.NavigateUrl = "/knowledge/ManagerMemberKnowledge_Question_Lp.aspx?MemberId=" & MemberId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
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
                sb.Append("<tr><th scope=""col"">狀態</th><th scope=""col"">問題標題</th><th scope=""col"">活動得分</th><th scope=""col"">知識分類</th><th scope=""col"">上架日期</th><th scope=""col"">建立日期</th><th scope=""col"">更新日期</th>")
                sb.Append("<th scope=""col"">討論</th><th scope=""col"">瀏覽</th><th scope=""col"">專家</th></tr>")

				Dim totalDiscuss As String
                For i = 0 To PageSize - 1

                    sb.Append("<tr>")
                    
					'發問狀態：正常 不公開 刪除
					IF myDataRow("Status") = "N" Then '正常
						If myDataRow("fCTUPublic") = "N" Then
							sb.Append("<td>不公開</td>")
						Else
							sb.Append("<td>&nbsp;</td>")
						End If
					Else
						sb.Append("<td>已刪除</td>")
					End If
					
					link = ""
					link = "<a href=""http://" & kmwebsysSite & "/KnowledgeForum/cp_question.asp?activity=Y&iCUItem=" & myDataRow("iCUItem") & "&MemberId=" & MemberId & "&BackActivePageNumber=" & PageNumber & "&BackActivePageSize=" & PageSize & """ target=""_blank"">"
					link &= myDataRow("sTitle") + "</a>"
					
                    sb.Append("<td>" & link & "</td>")
                    sb.Append("<td>" & myDataRow("Grade") & "</td>")
                    sb.Append("<td>" & myDataRow("mValue") & "</td>")
                    sb.Append("<td>[" & Date.Parse(myDataRow("CreateTime")).ToShortDateString & "]</td>")
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
