Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration
Imports System.Web
Imports System.IO



Partial Class KnowledgeActivityRankDetail
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'debug

        '---list---
        Dim PageNumber As Integer
        Dim PageSize As Integer

        If Not Page.IsPostBack Then
            Try
                If Request.QueryString("PageSize") = "" Then
                    PageSize = 15
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
                Response.Redirect("/")
                Response.End()
            End Try
        Else
            PageSize = PageSizeDDL.SelectedValue
            PageNumber = PageNumberDDL.SelectedValue
        End If
		
		If  Request.QueryString("ScoreS") <> "" Then
			TextBoxScoreS.Text = Request.QueryString("ScoreS")
		End If
		
		If  Request.QueryString("ScoreE") <> "" Then
			TextBoxScoreE.Text = Request.QueryString("ScoreE")
		End If
		
		If  Request.QueryString("QueryMember") <> "" Then
			TextBoxMember.Text = Request.QueryString("QueryMember")
		End If
		'Response.Write(TextBoxScoreE.Text)
		'上一頁 下一頁url參數
		Dim qryStr As String
        '組合篩選條件
        'Dim condition As String
        If TextBoxMember.Text.Trim() <> "" Then
            'condition = condition & "AND ( Total.MemberId LIKE @MemberId OR (M.realname LIKE @MemberId OR M.realname LIKE @MemberIdEncode)OR (M.nickname LIKE @MemberId OR M.nickname LIKE @MemberIdEncode)) "
			qryStr = "&QueryMember=" & TextBoxMember.Text
        End If

        If TextBoxScoreS.Text.Trim() <> "" And TextBoxScoreE.Text.Trim() <> "" Then
            'condition = condition & "AND Grade Between @ScoreS AND @ScoreE "
			qryStr &= "&ScoreS=" & TextBoxScoreS.Text & "&ScoreE=" & TextBoxScoreE.Text

        End If

        Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ConnString").ConnectionString


        Dim sqlString As StringBuilder = New StringBuilder()
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
            myConnection.Open()

            'sqlString.Append("SELECT KA.MemberId,M.nickname,M.realname,SUM(KA.Grade) AS Grade FROM dbo.KnowledgeActivity KA ")
            'sqlString.Append("INNER JOIN dbo.Member M ON KA.MemberId = M.account ")
            'sqlString.Append("WHERE State <> 0 GROUP BY KA.MemberId,M.nickname,M.realname ORDER BY KA.Grade DESC")


            sqlString.Append("SELECT Total.MemberId,M.nickname,M.realname,SUM(Total.Grade) As Grade FROM ")
            sqlString.Append("(	SELECT KA.MemberId,SUM(KA.Grade) AS Grade FROM dbo.KnowledgeActivity KA ")
            sqlString.Append("WHERE type BETWEEN 1 AND 2 ")
            sqlString.Append("GROUP BY KA.MemberId ")
            sqlString.Append("UNION ALL ")
            sqlString.Append("SELECT Temp.MemberId,SUM(temp.Grade) AS Grade ")
            sqlString.Append("FROM( SELECT KA.MemberId,CASE  WHEN SUM(KA.Grade)>4 THEN 4 ELSE SUM(KA.Grade) END AS Grade  ")
            sqlString.Append("FROM dbo.KnowledgeActivity KA ")
            sqlString.Append("INNER JOIN dbo.KnowledgeForum KF ON ka.CUItemId = KF.gicuitem ")
            sqlString.Append("WHERE KA.State =1 AND KA.type BETWEEN 3 AND 4 AND KF.Status = 'N' ")
            sqlString.Append("GROUP BY KF.ParentIcuitem, KA.MemberId) AS Temp ")
            sqlString.Append("GROUP BY Temp.MemberId) Total ")
            sqlString.Append("INNER JOIN dbo.Member M ON Total.MemberId = M.account ")
			sqlString.Append("WHERE M.status <> 'N'")

            If TextBoxMember.Text.Trim() <> "" Then
                sqlString.Append("AND ( Total.MemberId LIKE @MemberId OR (M.realname LIKE @MemberId OR M.realname LIKE @MemberIdEncode)OR (M.nickname LIKE @MemberId OR M.nickname LIKE @MemberIdEncode)) ")
            End If

            sqlString.Append("GROUP BY Total.MemberId,M.nickname,M.realname ")

            If TextBoxScoreS.Text.Trim() <> "" And TextBoxScoreE.Text.Trim() <> "" Then
                sqlString.Append("HAVING SUM(Total.Grade) Between @ScoreS AND @ScoreE ")
            End If
            sqlString.Append("ORDER BY Grade desc")



            myCommand = New SqlCommand(sqlString.ToString(), myConnection)

            If TextBoxMember.Text.Trim() <> "" Then
				Dim txt As String = GetEncodeString(TextBoxMember.Text)
                myCommand.Parameters.AddWithValue("@MemberId", "%" & TextBoxMember.Text.Trim() & "%")
                myCommand.Parameters.AddWithValue("@MemberIdEncode", "%" & txt & "%")
            End If

            If TextBoxScoreS.Text.Trim() <> "" And TextBoxScoreE.Text.Trim() <> "" Then
                myCommand.Parameters.AddWithValue("@ScoreS", TextBoxScoreS.Text.Trim())
                myCommand.Parameters.AddWithValue("@ScoreE", TextBoxScoreE.Text.Trim())
            End If

'Response.Write(sqlString )

            'linkExport.NavigateUrl = "KnowledgeActivityRankExport.aspx?MemberId=" & TextBoxMember.Text.Trim() & "&ScoreS=" & TextBoxScoreS.Text.Trim() & "&ScoreE=" & TextBoxScoreE.Text.Trim()
			'linkExport.NavigateUrl = "KnowledgeRankDetailExport.asp?MemberId=" & TextBoxMember.Text.Trim() & "&MemberIdEncode=" & GetEncodeString(TextBoxMember.Text) & "&ScoreS=" & TextBoxScoreS.Text.Trim() & "&ScoreE=" & TextBoxScoreE.Text.Trim()
			
			Dim kmwebsysSite As String = WebConfigurationManager.AppSettings("kmwebsysSite") 
			linkExport.NavigateUrl = "http://" & kmwebsysSite & "/KnowledgeForum/KnowledgeRankDetailExport.asp?MemberId=" & TextBoxMember.Text.Trim() & "&ScoreS=" & TextBoxScoreS.Text.Trim() & "&ScoreE=" & TextBoxScoreE.Text.Trim()
            myReader = myCommand.ExecuteReader()

            myTable = New DataTable()
            myTable.Load(myReader)

            Total = myTable.Rows().Count()
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myCommand.Dispose()

            If Total <> 0 Then

                PageCount = Int(Total / PageSize + 0.999)

                If PageCount < PageNumber Then
                    PageNumber = PageCount
                End If

                PageNumberText.Text = PageNumber.ToString()
                TotalPageText.Text = PageCount.ToString()
                TotalRecordText.Text = Total.ToString()
                If PageSize = 15 Then
                    PageSizeDDL.SelectedIndex = 0
                ElseIf PageSize = 30 Then
                    PageSizeDDL.SelectedIndex = 1
                ElseIf PageSize = 50 Then
                    PageSizeDDL.SelectedIndex = 2
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
                    PreviousLink.NavigateUrl = "KnowledgeActivityRankDetail.aspx?PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize & qryStr
                Else
                    PreviousText.Enabled = False
                End If
                If PageNumberDDL.SelectedValue < PageCount Then
                    NextLink.NavigateUrl = "KnowledgeActivityRankDetail.aspx?PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize & qryStr
                Else
                    NextText.Enabled = False
                End If

                Dim i As Integer = 0
                Dim link As String = ""
                Dim sb As New StringBuilder


                sb.Append("<table class=""type02"" summary=""結果條列式"" width=""100%"">")
                sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                sb.Append("<th scope=""col"" width=""10%"">帳號</th>")
                sb.Append("<th scope=""col"" width=""10%"">姓名</th>")
                sb.Append("<th scope=""col"" width=""10%"">暱稱</th>")
                sb.Append("<th scope=""col"" width=""10%"">活動得分</th>")
                sb.Append("<th scope=""col"" width=""15%"">已達最低抽獎資格</th>")
                sb.Append("<th scope=""col"" width=""10%"">可抽獎次數</th>")
                sb.Append("<th scope=""col"" width=""30%"">明細</th>")
                sb.Append("</tr>")

                For i = 0 To PageSize - 1

                    link = "<a href=""/knowledge/ManagerMemberKnowledge_Question_Lp.aspx?MemberId=" & myDataRow("memberId") & "&BackPageNumber=" & PageNumber & "&BackPageSize=" & PageSize & " ""> 活動發問 </a> &nbsp;&nbsp;   "
                    link &= "<a href=""/knowledge/ManagerMemberKnowledge_Discuss_Lp.aspx?MemberId=" & myDataRow("memberId") & "&BackPageNumber=" & PageNumber & "&BackPageSize=" & PageSize & """> 活動討論 </a>"
                    sb.Append("<tr><td>" & (PageSize * (PageNumber - 1)) + i + 1 & "</td>")
					
					'帳號
					If myDataRow("MemberId") IsNot DBNull.Value Then
                        If myDataRow("MemberId").ToString().Trim <> "" Then
                           sb.Append("<td>" & myDataRow("MemberId") & "</td>")
                        Else
                            sb.Append("<td>&nbsp;</td>")
                        End If
					Else
						sb.Append("<td>&nbsp;</td>")
					End If
					
					'姓名
					If myDataRow("REALNAME") IsNot DBNull.Value Then
                        If myDataRow("REALNAME").ToString().Trim <> "" Then
                            sb.Append("<td>" & myDataRow("REALNAME") & "</td>")
                        Else
                            sb.Append("<td>&nbsp;</td>")
                        End If
                    Else
                        sb.Append("<td>&nbsp;</td>")
                    End If
					
					'暱稱
                    If myDataRow("NICKNAME") IsNot DBNull.Value Then
                        If myDataRow("NICKNAME").ToString().Trim <> "" Then
                            sb.Append("<td>" & myDataRow("NICKNAME") & "</td>")
                        Else
                            sb.Append("<td>&nbsp;</td>")
                        End If
                    Else
                        sb.Append("<td>&nbsp;</td>")
                    End If

                    sb.Append("<td align=""center"">" & myDataRow("Grade").ToString & "</td>")
                    sb.Append("<td align=""center"">" & CheckCertification(CType(myDataRow("Grade"), Integer)) & "</td>")
                    sb.Append("<td align=""center"">" & GetAwardNum(CType(myDataRow("Grade"), Integer)) & "</td>")
                    sb.Append("<td align=""center"">" & link & "</td>")
                    sb.Append("</tr>")

                    Position += 1
                    If myTable.Rows.Count <= Position Then
                        Exit For
                    End If
                    myDataRow = myTable.Rows(Position)
                Next

                sb.Append("</table>")

                TableText.Text = sb.ToString()

            Else

                Dim sb As New StringBuilder

                PageNumberText.Text = "0"
                TotalPageText.Text = "0"
                TotalRecordText.Text = "0"

                sb.Append("<table width=""100%"" summary=""結果條列式"">")
                sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                sb.Append("<th scope=""col"" width=""10%"">帳號</th>")
				sb.Append("<th scope=""col"" width=""10%"">姓名</th>")
                sb.Append("<th scope=""col"" width=""10%"">暱稱</th>")
                sb.Append("<th scope=""col"" width=""10%"">活動得分</th>")
                sb.Append("<th scope=""col"" width=""15%"">已達最低抽獎資格</th>")
                sb.Append("<th scope=""col"" width=""10%"">可抽獎次數</th>")
                sb.Append("<th scope=""col"" width=""30%"">明細</th>")
                sb.Append("</tr>")
                sb.Append("</table>")

                TableText.Text = sb.ToString()

            End If



        Catch ex As Exception
            If Request("debug") = "true" Then
                Response.Write(ex.ToString())
                Response.End()
            End If
            Response.Write(ex.Message.ToString())
            'Response.Redirect("/")
            Response.End()
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try

    End Sub
	

    Function CheckCertification(ByVal grade As Integer) As String
        ' 積分達10分以上即有抽獎資格

        If grade >= 10 Then
            Return "☆"
        Else
            Return "&nbsp;"
        End If
    End Function
	
	Function GetEncodeString(ByVal orignalStr) As String
		dim old AS String
		dim new_w AS String
		dim iStr AS Integer
		old = orignalStr
		new_w = ""
		for iStr = 1 to orignalStr.length
			new_w = new_w & "&#" & ascw(mid(old,iStr,1))& ";"
		next
		
		return new_w
	End Function

    Function GetAwardNum(ByVal grade As Integer) As String
        
		' 積分達10分以上即可參加抽獎
		Dim num As string 
		If grade < 10 Then
			num = "0"
		ElseIf grade >= 10 And grade < 20 Then
            num = "1"
        ElseIf grade >= 20 And grade < 30 Then
            num = "2"
        Elseif grade >= 30 And grade < 40 Then
            num = "3"
		Elseif grade >= 40 And grade < 50 Then
            num = "4"
        Elseif grade >= 50 Then
            num = "5"			
		End If

        Return num

    End Function



End Class
