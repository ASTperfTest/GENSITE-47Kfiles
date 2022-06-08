Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Top
  Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

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

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("fishBowlConnString").ConnectionString
    Dim sqlString as StringBuilder= new StringBuilder()
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

		sqlString.Append(" SELECT a.account_id,case when e.stars is NULL then 0 else e.stars end as STAR,b.realname as REALNAME,b.nickname as NICKNAME, " & ControlChars.CrLf)
		sqlString.Append(" case when d.fish_score is NULL then 0 " & ControlChars.CrLf)
		sqlString.Append(" else fish_score " & ControlChars.CrLf)
		sqlString.Append(" end as SCORE, " & ControlChars.CrLf)
		sqlString.Append(" case when f.SUCCESS_FISH_COUNT is NULL then 0 " & ControlChars.CrLf)
		sqlString.Append(" else f.SUCCESS_FISH_COUNT " & ControlChars.CrLf)
		sqlString.Append(" end as SUCCESS_FISH_COUNT " & ControlChars.CrLf)
		sqlString.Append(" FROM fishbowl_environment a  " & ControlChars.CrLf)
		sqlString.Append(" LEFT JOIN fishbowl_account b " & ControlChars.CrLf)
		sqlString.Append(" on a.account_id=b.id " & ControlChars.CrLf)
		sqlString.Append(" LEFT JOIN ( " & ControlChars.CrLf)
		sqlString.Append(" SELECT SUM(temp_score) as fish_score,account_id FROM " & ControlChars.CrLf)
		sqlString.Append(" (  " & ControlChars.CrLf)
		sqlString.Append(" 	SELECT (z.health_score + z.full_score + z.water_score) AS temp_score,y.account_id  " & ControlChars.CrLf)
		sqlString.Append(" 	FROM fishbowl_pets_score z " & ControlChars.CrLf)
		sqlString.Append(" 	LEFT JOIN fishbowl_pets y " & ControlChars.CrLf)
		sqlString.Append(" 	on z.fish_id=y.id " & ControlChars.CrLf)
		sqlString.Append(" ) as u  " & ControlChars.CrLf)
		sqlString.Append(" GROUP BY account_id " & ControlChars.CrLf)
		sqlString.Append(" ) d " & ControlChars.CrLf)
		sqlString.Append(" on a.account_id=d.account_id " & ControlChars.CrLf)
		sqlString.Append(" LEFT JOIN ( " & ControlChars.CrLf)
		sqlString.Append(" SELECT account_id,SUM(stars) as stars " & ControlChars.CrLf)
		sqlString.Append(" FROM fishbowl_pets " & ControlChars.CrLf) 
		sqlString.Append(" GROUP BY account_id " & ControlChars.CrLf)
		sqlString.Append(" ) e " & ControlChars.CrLf)
		sqlString.Append(" on a.account_id=e.account_id " & ControlChars.CrLf)
		sqlString.Append(" LEFT JOIN ( " & ControlChars.CrLf)
		sqlString.Append(" SELECT count(distinct fish_type) as SUCCESS_FISH_COUNT,account_id  " & ControlChars.CrLf)
		sqlString.Append(" FROM fishbowl_pets " & ControlChars.CrLf) 
		sqlString.Append(" WHERE status=2 and is_active=1 " & ControlChars.CrLf)
		sqlString.Append(" GROUP BY account_id " & ControlChars.CrLf)
		sqlString.Append(" ) f " & ControlChars.CrLf)
		sqlString.Append(" on a.account_id=f.account_id " & ControlChars.CrLf)
		sqlString.Append(" order by SUCCESS_FISH_COUNT DESC,STAR DESC, SCORE DESC" & ControlChars.CrLf)
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
          PreviousLink.NavigateUrl = "Top.aspx?PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Enabled = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextLink.NavigateUrl = "Top.aspx?PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Enabled = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""
        Dim sb As New StringBuilder

        sb.Append("<table class=""type02"" summary=""結果條列式"">")
        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
        sb.Append("<th scope=""col"" width=""20%"">姓名｜暱稱</th>")
        sb.Append("<th scope=""col"" width=""20%"">已成功養殖品項數</th>")
        sb.Append("<th scope=""col"" width=""20%"">星級</th>")
        sb.Append("<th scope=""col"" width=""20%"">分數</th>")
        sb.Append("<th scope=""col"" width=""20%"">已達最低抽獎資格</th>")
        sb.Append("</tr>")

        For i = 0 To PageSize - 1

          sb.Append("<tr><td>" & (PageSize * (PageNumber - 1)) + i + 1 & "</td>")

          If myDataRow("NICKNAME") IsNot DBNull.Value Then
            If myDataRow("NICKNAME") <> "" Then
              sb.Append("<td>" & myDataRow("NICKNAME") & "</td>")
            Else
              sb.Append("<td>" & myDataRow("REALNAME").ToString.Substring(0, 1) & "＊" & myDataRow("REALNAME").ToString.Substring(myDataRow("REALNAME").ToString.Trim.Length - 1, 1) & "</td>")
            End If
          ElseIf myDataRow("REALNAME") IsNot DBNull.Value Then
            If myDataRow("REALNAME") <> "" Then
              sb.Append("<td>" & (myDataRow("REALNAME").ToString.Substring(0, 1) & "＊" & myDataRow("REALNAME").ToString.Substring(myDataRow("REALNAME").ToString.Trim.Length - 1, 1)) & "</td>")
            End If
          Else
            sb.Append("<td>&nbsp;</td>")
          End If
          sb.Append("<td align=""center"">" & myDataRow("SUCCESS_FISH_COUNT").ToString & "</td>")
          sb.Append("<td align=""center"">" & myDataRow("STAR").ToString & "</td>")
		  sb.Append("<td align=""center"">" & myDataRow("SCORE").ToString & "</td>")

          sb.Append("<td align=""center"">" & CheckCertification(Ctype(myDataRow("STAR"),Integer),Ctype(myDataRow("SUCCESS_FISH_COUNT"),Integer)) & "</td>")

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

        sb.Append("<table class=""type02"" summary=""結果條列式"">")
        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
        sb.Append("<th scope=""col"" width=""20%"">姓名｜暱稱</th>")
        sb.Append("<th scope=""col"" width=""20%"">已成功種植作物數</th>")
        sb.Append("<th scope=""col"" width=""20%"">星級</th>")
        sb.Append("<th scope=""col"" width=""20%"">分數</th>")
        sb.Append("<th scope=""col"" width=""20%"">已達最低抽獎資格</th>")
        sb.Append("</tr>")
        sb.Append("</table>")

        TableText.Text = sb.ToString()

      End If

    Catch ex As Exception
      If Request("debug") = "true" Then
        Response.Write(ex.ToString())
        Response.End()
      End If
      Response.Redirect("/")
      Response.End()
    Finally
      If myConnection.State = ConnectionState.Open Then
        myConnection.Close()
      End If
    End Try

  End Sub

  Function CheckCertification(ByVal starCount As Integer,ByVal successFishCount As Integer) As String
      If starCount >= 10 AndAlso successFishCount >=1 Then
          Return "☆"
      Else
		  Return "&nbsp;"	
      End If
  End Function

End Class
