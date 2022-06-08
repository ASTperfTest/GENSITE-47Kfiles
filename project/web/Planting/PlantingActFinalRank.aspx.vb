Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Planting_PlantingActRank
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

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("PlantingConnString").ConnectionString
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
      myConnection.Open()

      'sqlString = "SELECT DISTINCT dbo.LBG_GUEST.UID, dbo.LBG_GUEST.SCORE, dbo.LBG_GUEST.NICKNAME, dbo.LBG_GUEST.REALNAME, dbo.LBG_GUEST.STAR "
      'sqlString &= "FROM dbo.LBG_GUEST INNER JOIN dbo.LBG_GUEST_N2M ON dbo.LBG_GUEST.UID = dbo.LBG_GUEST_N2M.FROMID "
      'sqlString &= "ORDER BY dbo.LBG_GUEST.SCORE DESC, dbo.LBG_GUEST.STAR DESC "

      sqlString = "SELECT UID, TOTALSCORE as SCORE, TOTALSTAR AS STAR ,"
      sqlString &= " NICKNAME, REALNAME "
      sqlString &= " from LBG_SCORE_SUMMARY_EX order by TOTALSTAR DESC, TOTALSCORE DESC "


      myCommand = New SqlCommand(sqlString, myConnection)
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
          PreviousLink.NavigateUrl = "/Planting/PlantingActFinalRank.aspx?PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Enabled = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextLink.NavigateUrl = "/Planting/PlantingActFinalRank.aspx?PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Enabled = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""
        Dim sb As New StringBuilder

        sb.Append("<table class=""type02"" summary=""結果條列式"">")
        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th>")
        sb.Append("<th scope=""col"" width=""20%"">姓名｜暱稱</th>")
        sb.Append("<th scope=""col"" width=""20%"">分數</th>")
        sb.Append("<th scope=""col"" width=""20%"">星級</th>")
        sb.Append("<th scope=""col"" width=""20%"">已成功種植作物數</th>")
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
          sb.Append("<td align=""center"">" & myDataRow("SCORE").ToString & "</td>")
          sb.Append("<td align=""center"">" & myDataRow("STAR").ToString & "</td>")
          sb.Append("<td align=""center"">" & GetSuccessPlantCount(myDataRow("UID").ToString) & "</td>")
          sb.Append("<td align=""center"">" & CheckCertification(myDataRow("UID").ToString) & "</td>")

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
        sb.Append("<th scope=""col"" width=""20%"">分數</th>")
        sb.Append("<th scope=""col"" width=""20%"">星級</th>")
        sb.Append("<th scope=""col"" width=""20%"">已成功種植作物數</th>")
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

  Function GetSuccessPlantCount(ByVal uid As String) As String

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("PlantingConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim count As Integer = 0

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      'sqlString = "SELECT COUNT(DISTINCT dbo.LBG_SESSION.PLANTTYPE) as count FROM dbo.LBG_GUEST INNER JOIN dbo.LBG_GUEST_N2M "
      'sqlString &= "ON dbo.LBG_GUEST.UID = dbo.LBG_GUEST_N2M.FROMID INNER JOIN dbo.LBG_SESSION ON "
      'sqlString &= "dbo.LBG_GUEST_N2M.TOID = dbo.LBG_SESSION.UID WHERE (dbo.LBG_GUEST.UID = @uid)"

      sqlString = "select count(distinct PLANTTYPE) as VALID_PLANTTYPE  from LBG_SESSION s , LBG_GUEST_N2M n where n.FROMID in "
      sqlString &= " (select UID from LBG_GUEST g where g.UID=@uid "
      sqlString &= " ) and n.TOID = s.UID and (STATUS = 4 or STATUS = 2 or STATUS = 5)"



      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@uid", uid)
      count = myCommand.ExecuteScalar
      myCommand.Dispose()

    Catch ex As Exception

    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try

    Return count.ToString

  End Function

  Function CheckCertification(ByVal uid As String) As String

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("PlantingConnString").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim flag As Boolean = False

    myConnection = New SqlConnection(ConnString)
    Try
      myConnection.Open()
      'sqlString = "SELECT COUNT(DISTINCT dbo.LBG_SESSION.PLANTTYPE) AS PlantCount, dbo.LBG_GUEST.STAR "
      'sqlString &= "FROM dbo.LBG_GUEST INNER JOIN dbo.LBG_GUEST_N2M ON dbo.LBG_GUEST.UID = dbo.LBG_GUEST_N2M.FROMID "
      'sqlString &= "INNER JOIN dbo.LBG_SESSION ON dbo.LBG_GUEST_N2M.TOID = dbo.LBG_SESSION.UID "
      'sqlString &= "WHERE (dbo.LBG_GUEST.UID = @uid) GROUP BY dbo.LBG_GUEST.STAR"

      sqlString = "select count(distinct PLANTTYPE) as PlantCount from LBG_SESSION s , LBG_GUEST_N2M n where n.FROMID in "
      sqlString &= " (select UID from LBG_GUEST g where g.UID=@uid "
      sqlString &= " and g.STAR >= 6 ) and n.TOID = s.UID and (STATUS = 4  or STATUS = 2 or STATUS = 5) "




      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@uid", uid)
      myReader = myCommand.ExecuteReader
      If myReader.Read Then
        'If CType(myReader("PlantCount"), Integer) >= 2 And CType(myReader("STAR"), Integer) >= 6 Then
        '  flag = True
        'End If
        If CType(myReader("PlantCount"), Integer) >= 2 Then
          flag = True
        End If
      End If
      myCommand.Dispose()

    Catch ex As Exception
    Finally
      If myConnection.State = ConnectionState.Open Then myConnection.Close()
    End Try
    If flag Then
      Return "☆"
    Else
      Return "&nbsp;"
    End If

  End Function

End Class
