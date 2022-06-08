Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_myknowledge_pedia
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim type As String = "A"
        Dim Keyword As String
        Dim sb As StringBuilder
        Dim MemberId As String
        Dim Script As String = "<script>alert('連線逾時或尚未登入，請登入會員');window.location.href='/knowledge/knowledge.aspx';</script>"

        '-----------------------------------------------------------------------------
        If Session("memID") = "" Or Session("memID") Is Nothing Then
            MemberId = ""
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "NoId", Script)
            Exit Sub
        Else
            MemberId = Session("memID").ToString()
        End If
        '-----------------------------------------------------------------------------

        If Request.QueryString("type") <> "A" And Request.QueryString("type") <> "B" Then
            type = "A"
        Else
            type = Request.QueryString("type")
        End If

        If Request.QueryString("keyword") IsNot Nothing Then
            Keyword = Web.HttpUtility.UrlDecode(Request.QueryString("keyword"))
        Else
            Keyword = ""
        End If
        '-----------------------------------------------------------------------------
        TabText.Text = WebUtility.GetMyAreaLinks(MyAreaLink.myknowledge_pedia)

        '-----------------------------------------------------------------------------
        Dim linkA As String = "/Pedia/PediaContent.aspx?AId={0}"
        Dim linkB As String = "/Pedia/PediaExplainContent.aspx?AId={1}&PAId={0}"
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
                Response.Redirect("/knowledge/myknowledge_pedia.aspx")
                Response.End()
            End Try
        Else
            PageSize = PageSizeDDL.SelectedValue
            If PageNumberDDL.SelectedValue = "" Then
                PageNumber = 1
            Else
                PageNumber = PageNumberDDL.SelectedValue
            End If
            Keyword = ""
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
            sqlString = "SELECT CuDTGeneric.xPostDate, CuDTGeneric.ClickCount, Pedia.memberId, Pedia.commendTime, Pedia.quoteIcuitem, Pedia.parentIcuItem, "
            sqlString &= "CuDTGeneric.sTitle, CuDTGeneric.iCUItem FROM CuDTGeneric INNER JOIN Pedia ON CuDTGeneric.iCUItem = Pedia.gicuitem "
            sqlString &= "WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (Pedia.xStatus = 'Y') AND (Pedia.memberId = @memberid) AND (CuDTGeneric.iCTUnit = @ictunit) "
            If Keyword <> "" And Keyword <> "請輸入關鍵字" Then
                sqlString &= "AND CuDTGeneric.sTitle LIKE '%" & Keyword & "%' "
            End If
            sqlString &= "ORDER BY CuDTGeneric.xPostDate DESC"

            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
            myCommand.Parameters.AddWithValue("@memberid", MemberId)
            If type = "A" Then
                myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("PediaCtUnitId"))
            Else
                myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("PediaAdditionalCtUnitId"))
            End If

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
                    PreviousLink.NavigateUrl = "/knowledge/myknowledge_pedia.aspx?type=" & type & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                Else
                    PreviousText.Visible = False
                    PreviousImg.Visible = False
                End If
                If PageNumberDDL.SelectedValue < PageCount Then
                    NextText.Visible = True
                    NextImg.Visible = True
                    NextLink.NavigateUrl = "/knowledge/myknowledge_pedia.aspx?type=" & type & "&Keyword=" & Web.HttpUtility.UrlEncode(Keyword) & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                Else
                    NextText.Visible = False
                    NextImg.Visible = False
                End If

                Dim i As Integer = 0
                Dim link As String = ""

                sb = New StringBuilder

                sb.Append("<div class=""lp"">")
                sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"">")
                sb.Append("<tr><th scope=""col"">詞彙標題</th><th scope=""col"">建立日期</th><th scope=""col"">更新日期</th><th scope=""col"">瀏覽數</th></tr>")

                For i = 0 To PageSize - 1

                    sb.Append("<tr>")
                    If type = "A" Then
                        link = linkA.Replace("{0}", myDataRow("iCuitem"))
                    Else
                        link = linkB.Replace("{1}", myDataRow("iCuitem")).Replace("{0}", myDataRow("parentIcuItem"))
                    End If
                    sb.Append("<td><a href=""" & link & """ target=""_blank"">" & myDataRow("sTitle") & "</a></td>")
                    sb.Append("<td>[" & Date.Parse(myDataRow("commendTime")).ToShortDateString & "]</td>")
                    sb.Append("<td>[" & Date.Parse(myDataRow("xPostDate")).ToShortDateString & "]</td>")
                    sb.Append("<td>" & myDataRow("ClickCount") & "</td>")
                    sb.Append("</tr>")

                    Position += 1
                    If myTable.Rows.Count <= Position Then
                        Exit For
                    End If
                    myDataRow = myTable.Rows(Position)
                Next

                sb.Append("</table></div>")

                TableText.Text = sb.ToString()
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

    Protected Sub SearchBtn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SearchBtn.Click

        Dim keyword As String = Web.HttpUtility.UrlEncode(KeywordText.Text.Trim)

        Response.Redirect("/knowledge/myknowledge_pedia.aspx?keyword=" & keyword)

    End Sub

End Class
