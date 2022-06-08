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

    Dim ConnString As String = FarmDB.connString

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

            'sqlString.Append(" SELECT account.id, account.login_id, account.realname, account.nickname, game_data.top_level, game_data.top_score, game_data.top_timer ")
            'sqlString.Append(" FROM farm2009_account AS account INNER JOIN ")
            'sqlString.Append(" (SELECT account_id, MAX(max_level) AS top_level, MAX(total_score) AS top_score, MIN(total_timer) AS top_timer ")
            'sqlString.Append(" FROM ( ")
            'sqlString.Append(" SELECT account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer FROM farm2009_question_data GROUP BY account_id, current_round ")
            'sqlString.Append(" ) AS game_data_1 ")
            'sqlString.Append(" GROUP BY account_id ")
            'sqlString.Append(" ) AS game_data ON account.id = game_data.account_id ")
            'sqlString.Append(" ORDER BY game_data.top_level DESC, game_data.top_score DESC, game_data.top_timer")

            ' 2009.11.16 fixed by eddie
			sqlString.Append(" SELECT account.id, account.login_id, account.realname, account.nickname, game_data.top_level, game_data.top_score, game_data.top_timer ")
			sqlString.Append(" FROM farm2009_account account,  ")
			sqlString.Append(" ( ")
			sqlString.Append(" SELECT account_id, MAX(max_level) AS top_level, MAX(total_score) AS top_score, MIN(best_timer) AS top_timer  ")
			sqlString.Append(" FROM ")
			sqlString.Append(" ( ")
			sqlString.Append(" SELECT account_id, MAX(q_data.current_level) AS max_level, SUM(q_data.current_score) AS total_score, SUM(q_data.current_timer) AS total_timer,  ")
			sqlString.Append(" ( ")
			sqlString.Append(" SELECT TOP 1 total_timer FROM  ")
			sqlString.Append(" ( ")
			sqlString.Append(" SELECT account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer ")
			sqlString.Append(" FROM farm2009_question_data  ")
			sqlString.Append(" GROUP BY account_id, current_round ")
			sqlString.Append(" ) AS timer_table ")
			sqlString.Append(" WHERE (timer_table.account_id = q_data.account_id) AND (timer_table.max_level = max_level) AND (timer_table.total_score = total_score) ")
			sqlString.Append(" ) AS best_timer ")
			sqlString.Append(" FROM farm2009_question_data q_data ")
			sqlString.Append(" GROUP BY q_data.account_id, q_data.current_round ")
			sqlString.Append(" ) AS score_list ")
			sqlString.Append(" GROUP BY account_id ")
			sqlString.Append(" ) AS game_data ")
			sqlString.Append(" WHERE account.id = game_data.account_id ")
			sqlString.Append(" ORDER BY game_data.top_level DESC, game_data.top_score DESC, game_data.top_timer ")

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
                sb.Append("<th scope=""col"" width=""20%"">闖關數</th>")
                sb.Append("<th scope=""col"" width=""20%"">闖關得分</th>")
                sb.Append("<th scope=""col"" width=""20%"">闖關時間</th>")
                'sb.Append("<th scope=""col"" width=""20%"">已達最低抽獎資格</th>")
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
                    sb.Append("<td align=""center"">" & myDataRow("top_level").ToString & "</td>")
                    sb.Append("<td align=""center"">" & myDataRow("top_score").ToString & "</td>")

                    ' 2009.11.16 fixed by eddie
                    'sb.Append("<td align=""center"">" & getGameTimer(Convert.ToInt32(myDataRow("id")), Convert.ToInt32(myDataRow("top_score")), Convert.ToInt32(myDataRow("top_score")) & "</td>")
                    sb.Append("<td align=""center"">" & SecsToMins(Convert.ToInt32(myDataRow("top_timer"))) & "</td>")

                    'sb.Append("<td align=""center"">" & CheckCertification(CType(myDataRow("top_level"), Integer), CType(myDataRow("top_score"), Integer)) & "</td>")
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
                sb.Append("<th scope=""col"" width=""20%"">闖關數</th>")
                sb.Append("<th scope=""col"" width=""20%"">闖關得分</th>")
                sb.Append("<th scope=""col"" width=""20%"">闖關時間</th>")
                'sb.Append("<th scope=""col"" width=""20%"">已達最低抽獎資格</th>")
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

    Function CheckCertification(ByVal level As Integer, ByVal total_score As Integer) As String
        ' 尚未確定抽獎標準，暫定有玩完三關，且分數在100分以上
        If level >= 1 Then
            Return "☆"
        Else
            Return "&nbsp;"
        End If
    End Function

    Function getGameTimer(ByVal account_id As Integer, ByVal top_level As Integer, ByVal top_score As Integer) As String
        Dim top_timer As Integer = 0

        Dim conn As SqlConnection = FarmDB.createConnection()
        Dim cmd_all As SqlCommand = New SqlCommand()

        Using conn
            Using cmd_all
                Dim sb As StringBuilder = New StringBuilder()
                sb.Append("SELECT total_timer FROM (SELECT account_id, MAX(current_level) AS max_level, SUM(current_score) AS total_score, SUM(current_timer) AS total_timer")
                sb.Append(" FROM farm2009_question_data GROUP BY account_id, current_round) AS derivedtbl_1 WHERE (account_id = @ACCOUNT_ID) AND (max_level = @MAX_LEVEL) AND (total_score = @MAX_SCORE)")

                cmd_all.Connection = conn
                cmd_all.CommandText = sb.ToString()
                cmd_all.Parameters.Clear()
                cmd_all.Parameters.Add("@ACCOUNT_ID", SqlDbType.Int, 4).Value = account_id
                cmd_all.Parameters.Add("@MAX_LEVEL", SqlDbType.Int, 4).Value = top_level
                cmd_all.Parameters.Add("@MAX_SCORE", SqlDbType.Int, 4).Value = top_score

                conn.Open()

                Dim the_dr As SqlDataReader = cmd_all.ExecuteReader()
                Using the_dr
                    If the_dr.Read() = True Then
                        top_timer = Convert.ToInt32(the_dr("total_timer"))
                    End If
                End Using
                conn.Close()

            End Using
        End Using

        Dim mins As Integer = top_timer / 60
        Dim secs As Integer = top_timer Mod 60

        Return mins & "分" & secs & "秒"

    End Function

    ' 2009.11.16 added by eddie
    Function SecsToMins(ByVal top_timer As Integer) As String

        ' 2009.11.18 fixed
        Dim secs As Integer = top_timer Mod 60
        Dim mins As Integer = (top_timer - secs) / 60

        'Dim mins As Integer = top_timer / 60
        'Dim secs As Integer = top_timer Mod 60

        Return mins & "分" & secs & "秒"
    End Function

End Class
