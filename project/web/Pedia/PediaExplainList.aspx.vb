Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class Pedia_PediaExplainList
    Inherits System.Web.UI.Page


  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim paid As String

    paid = Request.QueryString("PAId")

    '---kpi use---reflash wont add grade---
    If Request.QueryString("kpi") <> "0" Then
      '---start of kpi user---20080911---vincent---
      If Session("memID") IsNot Nothing Or Session("memID") <> "" Then
        Dim browse As New KPIBrowse(Session("memID"), "browsePediaExplainLP", paid)
        browse.HandleBrowse()
      End If
      '---end of kpi use---
      Dim relink As String = "/Pedia/PediaExplainList.aspx?PAId=" & paid & "&kpi=0"
      Response.Redirect(relink)
      Response.End()
    End If
    '---end of kpi use---   

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
    Dim sql As String = ""
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
      sql = "SELECT sTitle FROM CuDTGeneric WHERE icuitem = @icuitem"
      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@icuitem", paid)
      myReader = myCommand.ExecuteReader()
      If myReader.Read Then
        WordTitle.Text = "<h3>詞目：" & myReader("sTitle") & "</h3>"
      Else
        Session("PediaErrMsg") = "詞目錯誤"
        ShowErrorMessage()
        Exit Sub
      End If
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()
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

    myConnection = New SqlConnection(ConnString)
    Try
      sql = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.xBody, Pedia.parentIcuItem, Member.realname, Member.nickname, "
      sql &= " Pedia.commendTime FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem INNER JOIN Member ON Pedia.memberId = Member.account "
      sql &= " WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (Pedia.parentIcuItem = @pid) ORDER BY Pedia.commendTime DESC "

      myConnection.Open()
      myCommand = New SqlCommand(sql, myConnection)
      myCommand.Parameters.AddWithValue("@pid", paid)
      myReader = myCommand.ExecuteReader()

      myTable = New DataTable()
      myTable.Load(myReader)
      If Not myReader.IsClosed Then
        myReader.Close()
      End If
      myCommand.Dispose()

      Total = myTable.Rows().Count()

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
          PreviousLink.NavigateUrl = "/Pedia/PediaExplainList.aspx?PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Enabled = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextLink.NavigateUrl = "/Pedia/PediaExplainList.aspx?PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Enabled = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""
        Dim sb As New StringBuilder
        Dim body As String = ""

        sb.Append("<table class=""type02"" summary=""結果條列式"">")
        sb.Append("<tr><th scope=""col"" width=""5%"">&nbsp;</th><th scope=""col"">補充解釋(摘要)</th><th scope=""col"" width=""10%"">發表者</th><th scope=""col"" width=""10%"">發佈日期</th></tr>")

        For i = 0 To PageSize - 1

          link = ""
          sb.Append("<tr><td>" & (PageSize * (PageNumber - 1)) + i + 1 & "</td>")

          If CType(myDataRow("xBody"), String).Length > 70 Then
            body = NoHTML(CType(myDataRow("xBody"), String).Substring(0, 70))
          Else
            body = NoHTML(CType(myDataRow("xBody"), String))
          End If

          link = "<a href=""/Pedia/PediaExplainContent.aspx?AId=" & myDataRow("iCUItem") & "&PAId=" & paid & """>" & body & "</a>"

          sb.Append("<td>" & link & "</td>")

          If myDataRow("nickname") IsNot DBNull.Value Then
            If myDataRow("nickname") <> "" Then
              sb.Append("<td>" & myDataRow("nickname") & "</td>")
            Else
              sb.Append("<td>" & myDataRow("realname").ToString.Substring(0, 1) & "＊" & myDataRow("realname").ToString.Substring(myDataRow("realname").ToString.Trim.Length - 1, 1) & "</td>")
            End If
          ElseIf myDataRow("realname") IsNot DBNull.Value Then
            If myDataRow("realname") <> "" Then
              sb.Append("<td>" & myDataRow("realname").ToString.Substring(0, 1) & "＊" & myDataRow("realname").ToString.Substring(myDataRow("realname").ToString.Trim.Length - 1, 1) & "</td>")
            Else
              sb.Append("<td>&nbsp;</td>")
            End If
          Else
            sb.Append("<td>&nbsp;</td>")
          End If

          sb.Append("<td>" & Date.Parse(myDataRow("commendTime")).ToShortDateString & "</td>")

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
                  "<tr><th scope=""col"">&nbsp;</th><th scope=""col"" width=""20%"">補充解釋(摘要)</th><th scope=""col"">發表者</th>" & _
                  "<th scope=""col"">發佈日期</th></tr></table>")


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

  Sub ShowErrorMessage()

    Dim str = "<script>alert('" & Session("PediaErrMsg") & "');window.location.href='/Pedia/PediaList.aspx';</script>"

    Session("PediaErrMsg") = ""

    Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "PediaError", str)

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
