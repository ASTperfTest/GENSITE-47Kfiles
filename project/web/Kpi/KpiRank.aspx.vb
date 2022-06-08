Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class KpiRank_KpiRank
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

    '---tab text--
    Dim cat As String = "A"
    If Request.QueryString("cat") IsNot Nothing And Request.QueryString("cat") <> "" Then      
      cat = Request.QueryString("cat")
    End If

    Dim str As String = "<li{A}><a href=""/KpiRank/KpiRank.aspx?cat=A"">總排行</a></li><li{B}><a href=""/KpiRank/KpiRank.aspx?cat=B"">知識排行</a></li>" & _
                        "<li{C}><a href=""/KpiRank/KpiRank.aspx?cat=C"">成就排行</a></li><li{D}><a href=""/KpiRank/KpiRank.aspx?cat=D"">貢獻排行</a></li>" & _
                        "<li{E}><a href=""/KpiRank/KpiRank.aspx?cat=E"">精神排行</a></li>"

    TabText.Text = "<ul class=""subjectcata"">"
    If cat = "A" Then
      TabText.Text &= str.Replace("{A}", " class=""active""").Replace("{B}", "").Replace("{C}", "").Replace("{D}", "").Replace("{E}", "")
    ElseIf cat = "B" Then
      TabText.Text &= str.Replace("{A}", "").Replace("{B}", " class=""active""").Replace("{C}", "").Replace("{D}", "").Replace("{E}", "")
    ElseIf cat = "C" Then
      TabText.Text &= str.Replace("{A}", "").Replace("{B}", "").Replace("{C}", " class=""active""").Replace("{D}", "").Replace("{E}", "")
    ElseIf cat = "D" Then
      TabText.Text &= str.Replace("{A}", "").Replace("{B}", "").Replace("{C}", "").Replace("{D}", " class=""active""").Replace("{E}", "")
    ElseIf cat = "E" Then
      TabText.Text &= str.Replace("{A}", "").Replace("{B}", "").Replace("{C}", "").Replace("{D}", "").Replace("{E}", " class=""active""")
    End If
    TabText.Text &= "</ul>"
      

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
      sqlString = " "

      If cat = "A" Then
        sqlString &= " SELECT MemberGradeSummary.browseTotal + MemberGradeSummary.loginTotal + MemberGradeSummary.shareTotal "
        sqlString &= " + MemberGradeSummary.contentTotal + MemberGradeSummary.additionalTotal AS total, "
        sqlString &= " Member.realname, Member.nickname FROM MemberGradeSummary INNER JOIN Member ON MemberGradeSummary.memberId = Member.account "
        sqlString &= " ORDER BY total DESC "
      ElseIf cat = "B" Then
        sqlString &= "SELECT MemberGradeSummary.browseTotal AS total, Member.realname, Member.nickname "
        sqlString &= "FROM MemberGradeSummary INNER JOIN Member ON MemberGradeSummary.memberId = Member.account "
        sqlString &= "ORDER BY total DESC "
      ElseIf cat = "C" Then
        sqlString &= "SELECT Member.realname, Member.nickname, MemberGradeSummary.contentTotal AS total "
        sqlString &= "FROM MemberGradeSummary INNER JOIN Member ON MemberGradeSummary.memberId = Member.account "
        sqlString &= "ORDER BY total DESC "
      ElseIf cat = "D" Then
        sqlString &= "SELECT Member.realname, Member.nickname, MemberGradeSummary.shareTotal AS total "
        sqlString &= "FROM MemberGradeSummary INNER JOIN Member ON MemberGradeSummary.memberId = Member.account "
        sqlString &= "ORDER BY total DESC "
      ElseIf cat = "E" Then
        sqlString &= "SELECT Member.realname, Member.nickname, MemberGradeSummary.loginTotal AS total "
        sqlString &= "FROM MemberGradeSummary INNER JOIN Member ON MemberGradeSummary.memberId = Member.account "
        sqlString &= "ORDER BY total DESC "
      End If

      myConnection.Open()
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
          PreviousLink.NavigateUrl = "/KpiRank/KpiRank.aspx?PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Enabled = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextLink.NavigateUrl = "/KpiRank/KpiRank.aspx?PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Enabled = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""
        Dim sb As New StringBuilder

        sb.Append(" <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"">")
        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th><th scope=""col"" width=""10%"">&nbsp;</th><th scope=""col"">等級</th>")
        sb.Append("<th scope=""col"">姓名｜暱稱</th><th scope=""col"">積分</th></tr>")

        For i = 0 To PageSize - 1

          sb.Append("<tr>")
          sb.Append("<td>" & (PageSize * (PageNumber - 1)) + i + 1 & "</td>")

          '---前三名加上圖片---
          If PageNumber = 1 Then
            If Position = 0 Then
                            sb.Append("<td><img src=""/xslGip/style3/images/1.gif"" alt=""第1名"" /></td>")
            ElseIf Position = 1 Then
                            sb.Append("<td><img src=""/xslGip/style3/images/2.gif"" alt=""第2名"" /></td>")
            ElseIf Position = 2 Then
                            sb.Append("<td><img src=""/xslGip/style3/images/3.gif"" alt=""第3名"" /></td>")
            Else
              sb.Append("<td>&nbsp;</td>")
            End If
          Else
            sb.Append("<td>&nbsp;</td>")
          End If
          If myDataRow("total") > 10000 Then
            sb.Append("<td>達人級會員</td>")
          ElseIf myDataRow("total") > 3500 Then
            sb.Append("<td>高手級會員</td>")
          ElseIf myDataRow("total") > 1000 Then
            sb.Append("<td>進階級會員</td>")
          Else
            sb.Append("<td>入門級會員</td>")
          End If
          If myDataRow("nickname") Is DBNull.Value Then
            sb.Append("<td>" & myDataRow("realname").ToString.Trim.Substring(0, 1) & "＊" & myDataRow("realname").ToString.Trim.Substring(myDataRow("realname").ToString.Trim.Length - 1, 1) & "</td>")
          Else
            If myDataRow("nickname") = "" Then
              sb.Append("<td>" & myDataRow("realname").ToString.Trim.Substring(0, 1) & "＊" & myDataRow("realname").ToString.Trim.Substring(myDataRow("realname").ToString.Trim.Length - 1, 1) & "</td>")
            Else
              sb.Append("<td>" & myDataRow("nickname").ToString & "</td>")
            End If
          End If

          sb.Append("<td>" & myDataRow("total") & "</td>")
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

        sb.Append(" <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"">")
        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th><th scope=""col"" width=""10%"">&nbsp;</th><th scope=""col"">等級</th>")
        sb.Append("<th scope=""col"">姓名｜暱稱</th><th scope=""col"">積分</th></tr></table>")

        TableText.Text = sb.ToString

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

End Class
