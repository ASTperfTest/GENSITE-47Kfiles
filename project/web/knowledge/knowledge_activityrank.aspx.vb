Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_activityrank
  Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Dim Activity As New ActivityFilter(System.Web.Configuration.WebConfigurationManager.AppSettings("ActivityId"), "")
    Dim ActivityFlag As Boolean = Activity.CheckActivity
	Dim sbTab As New StringBuilder
		
	If  ActivityFlag Then
		sbTab.Append("<a href=""/knowledge/knowledge_alp.aspx""><img src=""image/mn_5.gif"" width=""107"" height=""28"" /></a>")
	Else
		sbTab.Append("&nbsp;")
	End If
	
	TabText.Text = sbTab.tostring()
	
	
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
			sqlString.Append("GROUP BY Total.MemberId,M.nickname,M.realname ORDER BY Grade desc")

			

            myCommand = New SqlCommand(sqlString.ToString(), myConnection)
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
                    PreviousLink.NavigateUrl = "knowledge_activityrank.aspx?PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                Else
                    PreviousText.Enabled = False
                End If
                If PageNumberDDL.SelectedValue < PageCount Then
                    NextLink.NavigateUrl = "knowledge_activityrank.aspx?PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                Else
                    NextText.Enabled = False
                End If

                Dim i As Integer = 0
                Dim link As String = ""
                Dim sb As New StringBuilder
				
				sb.Append("<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""  id=""rank"">")
                sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                sb.Append("<th scope=""col"" width=""30%"">會員</th>")
                sb.Append("<th scope=""col"" width=""20%"">活動得分</th>")
				sb.Append("<th scope=""col"" width=""20%"">已達最低抽獎資格</th>")
                sb.Append("<th scope=""col"" width=""20%"">可抽獎次數</th>")             
                sb.Append("</tr>")

                For i = 0 To PageSize - 1

                    sb.Append("<tr><td>" & (PageSize * (PageNumber - 1)) + i + 1 & "</td>")

                    If myDataRow("NICKNAME") IsNot DBNull.Value Then
                        If myDataRow("NICKNAME") <> "" Then
                            sb.Append("<td>" & myDataRow("NICKNAME") & "</td>")
                        Else
							Dim name As String = myDataRow("realname").ToString.Trim
							  if instr(name,"&#")>0 and instr(name,";")>1 then
								name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
							  elseif instr(name,"&#")>0 and instr(name,";")<1 then
								name = name.Substring(0, 7) & "＊" & name.Substring(name.Length - 7, 7)
							  else
								name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
							  end if
							  sb.Append("<td>" & name & "</td>")		  
                            'sb.Append("<td>" & Trim(myDataRow("REALNAME")).ToString.Substring(0, 1) & "＊" & Trim(myDataRow("REALNAME")).ToString.Substring(Trim(myDataRow("REALNAME")).ToString.Trim.Length - 1, 1) & "</td>")
                        End If
                    ElseIf myDataRow("REALNAME") IsNot DBNull.Value Then
                        If myDataRow("REALNAME") <> "" Then
							Dim name As String = myDataRow("realname").ToString.Trim
							  if instr(name,"&#")>0 and instr(name,";")>1 then
								name = name.Substring(0, 8) & "＊" & name.Substring(name.Length - 8, 8)
							  elseif instr(name,"&#")>0 and instr(name,";")<1 then
								name = name.Substring(0, 7) & "＊" & name.Substring(name.Length - 7, 7)
							  else
								name = name.Substring(0, 1) & "＊" & name.Substring(name.Length - 1, 1)
							  end if
							  sb.Append("<td>" & name & "</td>")
                           ' sb.Append("<td>" & (Trim(myDataRow("REALNAME")).ToString.Substring(0, 1) & "＊" & Trim(myDataRow("REALNAME")).ToString.Substring(Trim(myDataRow("REALNAME")).ToString.Trim.Length - 1, 1)) & "</td>")
                        End If
                    Else
                        sb.Append("<td>&nbsp;</td>")
                    End If
					
					sb.Append("<td align=""center"">" & myDataRow("Grade").ToString & "</td>")
					sb.Append("<td align=""center"">" & CheckCertification(CType(myDataRow("Grade"), Integer)) & "</td>")
					sb.Append("<td align=""center"">" & GetAwardNum(CType(myDataRow("Grade"),Integer)) & "</td>")
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

				sb.Append("<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0""  id=""rank"">")
                sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
                sb.Append("<th scope=""col"" width=""30%"">會員</th>")
                sb.Append("<th scope=""col"" width=""30%"">活動得分</th>")
				sb.Append("<th scope=""col"" width=""20%"">已達最低抽獎資格</th>")
                sb.Append("<th scope=""col"" width=""30%"">抽獎次數</th>")             
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
		
		return num
		
	End Function
	
	
    

End Class
