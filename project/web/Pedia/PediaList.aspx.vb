Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Pedia_PediaList
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
      sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, CuDTGeneric.xBody, CuDTGeneric.vGroup "
      sqlString &= "FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
      sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') "

      If Request.QueryString("type") = "query" Then
        If Session("PediaQueryTitle") <> "" Then
          sqlString &= " AND (CuDTGeneric.sTitle LIKE '%" & Session("PediaQueryTitle") & "%') "
        End If
        If Session("PediaQueryEngTitle") <> "" Then
          sqlString &= " AND (Pedia.engTitle LIKE '%" & Session("PediaQueryEngTitle") & "%') "
        End If
        If Session("PediaQueryBody") <> "" Then
          sqlString &= " AND (CuDTGeneric.xBody LIKE '%" & Session("PediaQueryBody") & "%') "
        End If
        If Session("PediaQueryIsOpen") <> "" Then
          sqlString &= " AND (CuDTGeneric.vGroup = '" & Session("PediaQueryIsOpen") & "') "
        End If
      Else        
        If isOpenDDL.SelectedValue = "Y" Then
          sqlString &= "AND (CuDTGeneric.vGroup = 'Y') "
        ElseIf isOpenDDL.SelectedValue = "N" Then
          sqlString &= "AND (CuDTGeneric.vGroup = 'N') "
        Else
        End If
      End If

      sqlString &= "ORDER BY CuDTGeneric.xPostDate DESC "

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("PediaCtUnitId"))
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
          PreviousLink.NavigateUrl = "/Pedia/PediaList.aspx?PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Enabled = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextLink.NavigateUrl = "/Pedia/PediaList.aspx?PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Enabled = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""
        Dim sb As New StringBuilder

        sb.Append("<table class=""type02"" summary=""結果條列式"">" & _
                  "<tr><th scope=""col"" width=""5%"">&nbsp;</th><th scope=""col"" width=""20%"">詞目</th><th scope=""col"">名詞釋義</th>" & _
                  "<th scope=""col"" width=""12%"">發佈日期</th><th scope=""col"" width=""60"">開放補充</th></tr>")

        For i = 0 To PageSize - 1

          link = ""
          sb.Append("<tr><td align=""center"">" & (PageSize * (PageNumber - 1)) + i + 1 & "</td>")

          link = "<a href=""/Pedia/PediaContent.aspx?AId=" & myDataRow("iCUItem") & """>"
          link &= myDataRow.Item("sTitle") + "</a>"

          sb.Append("<td>" & link & "</td>")
          If CType(myDataRow.Item("xBody"), String).Length > 70 Then
            sb.Append("<td>" & NoHTML(CType(myDataRow.Item("xBody"), String).Substring(0, 70)) & "......</td>")
          Else
            sb.Append("<td>" & NoHTML(CType(myDataRow.Item("xBody"), String)) & "</td>")
          End If

          sb.Append("<td align=""center"">" & Date.Parse(myDataRow("xPostDate")).ToShortDateString & "</td>")
          If myDataRow("vGroup") = "Y" Then
            sb.Append("<td align=""center"">V</td>")
          Else
            sb.Append("<td align=""center"">&nbsp;</td>")
          End If

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

        sb.Append("<table class=""type02"" summary=""結果條列式"">" & _
                  "<tr><th scope=""col"">&nbsp;</th><th scope=""col"" width=""20%"">詞目</th><th scope=""col"">名詞釋義</th>" & _
                  "<th scope=""col"">發佈日期</th><th scope=""col"">開放補充</th></tr></table>")

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

  Function NoHTML(ByVal Htmlstring As String) As String

    Htmlstring = Regex.Replace(Htmlstring, "<script[^>]*?>.*?</script>", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "<(.[^>]*)>", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "([\r\n])[\s]+", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "-->", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "<!--.*", "", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(quot|#34);", "\", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(amp|#38);", "&", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(lt|#60);", "<", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(gt|#62);", ">", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(nbsp|#160);", "   ", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(iexcl|#161);", "\xa1", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(cent|#162);", "\xa2", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(pound|#163);", "\xa3", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&(copy|#169);", "\xa9", RegexOptions.IgnoreCase)
    Htmlstring = Regex.Replace(Htmlstring, "&#(\d+);", "", RegexOptions.IgnoreCase)
    Htmlstring = Htmlstring.Replace("<", "")
    Htmlstring = Htmlstring.Replace(">", "")
    Htmlstring = Htmlstring.Replace("\r\n", "")
    'Htmlstring = HttpContext.Current.Server.HtmlEncode(Htmlstring).Trim()
    Return Htmlstring

  End Function

End Class
