Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_alp
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Dim Activity As New ActivityFilter(WebConfigurationManager.AppSettings("ActivityId"), "")
    Dim ActivityFlag As Boolean = Activity.CheckActivity
	
	If not ActivityFlag Then 
		Response.Redirect("/knowledge/knowledge.aspx")
        Response.End()
	End If
	
    Dim ArticleType As String
    Dim CategoryId As String
    Dim PageNumber As Integer
    Dim PageSize As Integer

    If Request.QueryString("CategoryId") <> "A" And Request.QueryString("CategoryId") <> "B" And Request.QueryString("CategoryId") <> "C" _
        And Request.QueryString("CategoryId") <> "D" And Request.QueryString("CategoryId") <> "E" And Request.QueryString("CategoryId") <> "F" Then
      CategoryId = ""
    Else
      CategoryId = Request.QueryString("CategoryId")
    End If

    Dim sb As New StringBuilder

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
        Response.Redirect("/knowledge/knowledge.aspx")
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
      sqlString = "SELECT KnowledgeActivity.CreateTime,CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, KnowledgeForum.DiscussCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, "
      sqlString &= "(KnowledgeForum.GradeCount/(KnowledgeForum.GradePersonCount + 0.000001)) AS Grade, ISNULL(CuDTGeneric.vGroup,'') AS vGroup, CuDTGeneric.xNewWindow, "
      sqlString &= "CuDTGeneric.topCat, KnowledgeForum.BrowseCount, KnowledgeForum.HavePros FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
	  sqlString &= "INNER JOIN dbo.KnowledgeActivity ON CuDTGeneric.iCUItem = dbo.KnowledgeActivity.CUItemId "
      sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND ( dbo.CuDTGeneric.xNewWindow ='N') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N') AND ( CuDTGeneric.vGroup = 'A') "
	  sqlString &= " ORDER BY KnowledgeActivity.CreateTime DESC "

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
      myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
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
          PreviousLink.NavigateUrl = "/knowledge/knowledge_alp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Visible = False
          PreviousImg.Visible = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextText.Visible = True
          NextImg.Visible = True
          NextLink.NavigateUrl = "/knowledge/knowledge_alp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Visible = False
          NextImg.Visible = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""

        sb = New StringBuilder
        sb.Append("<div class=""lp"">")
        sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" id=""qa"">")
        sb.Append("<tr><th scope=""col"">&nbsp;</th><th scope=""col"">標題</th><th scope=""col"">上架日期</th><th scope=""col"">發問日期</th><th scope=""col"">討論數</th>")
        sb.Append("<th scope=""col"">瀏覽數</th>")

        sb.Append("</tr>")
		Dim totalDiscuss As Integer = 0
        For i = 0 To PageSize - 1

          link = ""
          sb.Append("<tr>")
          sb.Append("<td>" & (PageSize * (PageNumber - 1)) + i + 1 & ". </td>")

          link = "<a href=""/knowledge/knowledge_cp.aspx?ArticleId=" & myDataRow("iCUItem") & "&ArticleType=" & ArticleType & "&CategoryId=" & myDataRow("topCat") & """>"
          link &= myDataRow.Item("sTitle") + "</a>"

		  
		  '計算討論則數 
			totalDiscuss="0"
                    Dim knowledgeUtility As New KnowledgeUtility(myDataRow("iCUItem"))
                    totalDiscuss = knowledgeUtility.GetDiscussCount()
			'計算討論則數 end
                    

		  
		  
          sb.Append("<td align=""left"">" & link & "</td>")
          sb.Append("<td>[" & Date.Parse(myDataRow("CreateTime")).ToShortDateString & "]</td>")
		  sb.Append("<td>[" & Date.Parse(myDataRow("xPostDate")).ToShortDateString & "]</td>")
          'sb.Append("<td>" & myDataRow("DiscussCount") & "</td>")
                    sb.Append("<td>" & totalDiscuss.ToString & "</td>")
                    sb.Append("<td>" & myDataRow("BrowseCount") & "</td>")

          'If myDataRow("xNewWindow") = "Y" Then
            'sb.Append("<td>是</td>")
          'Else
           ' sb.Append("<td>&nbsp;</td>")
          'End If

          'If myDataRow("HavePros") = "Y" Then
          ' sb.Append("<td>V</td>")
          ' Else
          ' sb.Append("<td>&nbsp;</td>")
          ' End If

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
		sb = new StringBuilder

        PageNumberText.Text = "0"
		TotalPageText.Text = "0"
		TotalRecordText.Text = "0"

        sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" id=""qa"">")
        sb.Append("<tr><th scope=""col"">&nbsp;</th><th scope=""col"" width=""30%"">標題</th><th scope=""col"">上架日期</th><th scope=""col"">發問日期</th><th scope=""col"">討論數</th>")
        sb.Append("<th scope=""col"">瀏覽數</th>")
		sb.Append("</tr>")
		sb.Append("</table>")

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


End Class
