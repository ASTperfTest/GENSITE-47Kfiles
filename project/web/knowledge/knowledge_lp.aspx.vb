Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class knowledge_lp
    Inherits System.Web.UI.Page

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim ArticleType As String
    Dim CategoryId As String
    Dim PageNumber As Integer
    Dim PageSize As Integer
    Dim PathStr As String

    If Request.QueryString("ArticleType") <> "A" And Request.QueryString("ArticleType") <> "B" And Request.QueryString("ArticleType") <> "C" _
        And Request.QueryString("ArticleType") <> "D" And Request.QueryString("ArticleType") <> "E" And Request.QueryString("ArticleType") <> "F" Then
      ArticleType = "A"
    Else
      ArticleType = Request.QueryString("ArticleType")
    End If

    If Request.QueryString("CategoryId") <> "A" And Request.QueryString("CategoryId") <> "B" And Request.QueryString("CategoryId") <> "C" _
        And Request.QueryString("CategoryId") <> "D" And Request.QueryString("CategoryId") <> "E" And Request.QueryString("CategoryId") <> "F" Then
      CategoryId = ""
    Else
      CategoryId = Request.QueryString("CategoryId")
    End If

    PathStr = "&gt;<a href=""/knowledge/knowledge_lp.aspx?CategoryId=" & CategoryId & """>{0}</a>"
    If CategoryId = "A" Then
      PathStr = PathStr.Replace("{0}", "農")
    ElseIf CategoryId = "B" Then
      PathStr = PathStr.Replace("{0}", "林")
    ElseIf CategoryId = "C" Then
      PathStr = PathStr.Replace("{0}", "漁")
    ElseIf CategoryId = "D" Then
      PathStr = PathStr.Replace("{0}", "牧")
    ElseIf CategoryId = "E" Then
      PathStr = PathStr.Replace("{0}", "其他")
    ElseIf CategoryId = "F" Then
      PathStr = PathStr.Replace("{0}", "產銷經營管理系統")
    ElseIf CategoryId = "" Then
      PathStr = PathStr.Replace("{0}", "全部")
    End If
    '---page---
    Dim TabLink As String = "/knowledge/knowledge_lp.aspx?ArticleType={0}&CategoryId=" & CategoryId
    Dim sb As New StringBuilder

    sb.Append("<ul>")
    If ArticleType = "A" Then

      PathText.Text = PathStr & "&gt;<a href=""#"">最新發問</a>"

      sb.Append("<li class=""current""><a href=""" & TabLink.Replace("{0}", "A") & """><span>最新發問</span></a></li><li><a href=""" & TabLink.Replace("{0}", "B") & """><span>熱門討論</span></a></li><li><a href=""" & TabLink.Replace("{0}", "C") & """>")
      sb.Append("<span>最佳評價</span></a></li><li><a href=""" & TabLink.Replace("{0}", "D") & """><span>最多意見</span></a></li><li><a href=""" & TabLink.Replace("{0}", "E") & """><span>專家補充</span></a></li>")
      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "F") & """><span>推薦討論</span></a></li>")
    ElseIf ArticleType = "B" Then

      PathText.Text = PathStr & "&gt;<a href=""#"">熱門討論</a>"

      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "A") & """><span>最新發問</span></a></li><li class=""current""><a href=""" & TabLink.Replace("{0}", "B") & """><span>熱門討論</span></a></li><li><a href=""" & TabLink.Replace("{0}", "C") & """>")
      sb.Append("<span>最佳評價</span></a></li><li><a href=""" & TabLink.Replace("{0}", "D") & """><span>最多意見</span></a></li><li><a href=""" & TabLink.Replace("{0}", "E") & """><span>專家補充</span></a></li>")
      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "F") & """><span>推薦討論</span></a></li>")
    ElseIf ArticleType = "C" Then

      PathText.Text = PathStr & "&gt;<a href=""#"">最佳評價</a>"

      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "A") & """><span>最新發問</span></a></li><li><a href=""" & TabLink.Replace("{0}", "B") & """><span>熱門討論</span></a></li><li class=""current""><a href=""" & TabLink.Replace("{0}", "C") & """>")
      sb.Append("<span>最佳評價</span></a></li><li><a href=""" & TabLink.Replace("{0}", "D") & """><span>最多意見</span></a></li><li><a href=""" & TabLink.Replace("{0}", "E") & """><span>專家補充</span></a></li>")
      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "F") & """><span>推薦討論</span></a></li>")
    ElseIf ArticleType = "D" Then

      PathText.Text = PathStr & "&gt;<a href=""#"">最多意見</a>"

      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "A") & """><span>最新發問</span></a></li><li><a href=""" & TabLink.Replace("{0}", "B") & """><span>熱門討論</span></a></li><li><a href=""" & TabLink.Replace("{0}", "C") & """>")
      sb.Append("<span>最佳評價</span></a></li><li class=""current""><a href=""" & TabLink.Replace("{0}", "D") & """><span>最多意見</span></a></li><li><a href=""" & TabLink.Replace("{0}", "E") & """><span>專家補充</span></a></li>")
      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "F") & """><span>推薦討論</span></a></li>")
    ElseIf ArticleType = "E" Then

      PathText.Text = PathStr & "&gt;<a href=""#"">專家補充</a>"

      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "A") & """><span>最新發問</span></a></li><li><a href=""" & TabLink.Replace("{0}", "B") & """><span>熱門討論</span></a></li><li><a href=""" & TabLink.Replace("{0}", "C") & """>")
      sb.Append("<span>最佳評價</span></a></li><li><a href=""" & TabLink.Replace("{0}", "D") & """><span>最多意見</span></a></li><li class=""current""><a href=""" & TabLink.Replace("{0}", "E") & """><span>專家補充</span></a></li>")
      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "F") & """><span>推薦討論</span></a></li>")
    ElseIf ArticleType = "F" Then

      PathText.Text = PathStr & "&gt;<a href=""#"">推薦討論</a>"

      sb.Append("<li><a href=""" & TabLink.Replace("{0}", "A") & """><span>最新發問</span></a></li><li><a href=""" & TabLink.Replace("{0}", "B") & """><span>熱門討論</span></a></li><li><a href=""" & TabLink.Replace("{0}", "C") & """>")
      sb.Append("<span>最佳評價</span></a></li><li><a href=""" & TabLink.Replace("{0}", "D") & """><span>最多意見</span></a></li><li><a href=""" & TabLink.Replace("{0}", "E") & """><span>專家補充</span></a></li>")
      sb.Append("<li class=""current""><a href=""" & TabLink.Replace("{0}", "F") & """><span>推薦討論</span></a></li>")
    End If
    sb.Append("</ul>")
    TabText.Text = sb.ToString
    sb = Nothing

    '---list---

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

    Dim Activity As New ActivityFilter(WebConfigurationManager.AppSettings("ActivityId"), "")
    Dim ActivityFlag As Boolean = Activity.CheckActivity

    myConnection = New SqlConnection(ConnString)
    Try
      sqlString = "SELECT CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CuDTGeneric.xPostDate, KnowledgeForum.DiscussCount, KnowledgeForum.GradeCount, KnowledgeForum.GradePersonCount, "
      sqlString &= "(KnowledgeForum.GradeCount/(KnowledgeForum.GradePersonCount + 0.000001)) AS Grade, ISNULL(CuDTGeneric.vGroup,'') AS vGroup, CuDTGeneric.xNewWindow, "
      sqlString &= "CuDTGeneric.topCat, KnowledgeForum.BrowseCount, KnowledgeForum.HavePros FROM CuDTGeneric INNER JOIN KnowledgeForum ON CuDTGeneric.iCUItem = KnowledgeForum.gicuitem "
      sqlString &= "WHERE (CuDTGeneric.iCTUnit = @ictunit) AND (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = @siteid) AND (KnowledgeForum.Status = 'N') "

      '---農林漁牧其中一種---
      If CategoryId <> "" Then
                sqlString &= "AND (CuDTGeneric.topCat = @CategoryId ) "
      End If
      '---排序種類---
      If ArticleType = "A" Then
        sqlString &= "ORDER BY CuDTGeneric.xPostDate DESC"
      ElseIf ArticleType = "B" Then
        sqlString &= "ORDER BY KnowledgeForum.DiscussCount DESC"
      ElseIf ArticleType = "C" Then
        sqlString &= "ORDER BY Grade DESC"
      ElseIf ArticleType = "D" Then
        sqlString &= "ORDER BY KnowledgeForum.CommandCount DESC"
      ElseIf ArticleType = "E" Then
        sqlString &= "AND (KnowledgeForum.HavePros = 'Y') ORDER BY CuDTGeneric.xPostDate DESC"
      ElseIf ArticleType = "F" Then
        sqlString &= "AND (CuDTGeneric.xImportant <> 0) ORDER BY CuDTGeneric.xPostDate DESC, CuDTGeneric.xImportant DESC"
      End If

      myConnection.Open()
      myCommand = New SqlCommand(sqlString, myConnection)
      myCommand.Parameters.AddWithValue("@ictunit", WebConfigurationManager.AppSettings("KnowledgeQuestionCtUnitId"))
            myCommand.Parameters.AddWithValue("@siteid", WebConfigurationManager.AppSettings("SiteId"))
            myCommand.Parameters.AddWithValue("@CategoryId", CategoryId)
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
          PreviousLink.NavigateUrl = "/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
        Else
          PreviousText.Visible = False
          PreviousImg.Visible = False
        End If
        If PageNumberDDL.SelectedValue < PageCount Then
          NextText.Visible = True
          NextImg.Visible = True
          NextLink.NavigateUrl = "/knowledge/knowledge_lp.aspx?ArticleType=" & ArticleType & "&CategoryId=" & CategoryId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
        Else
          NextText.Visible = False
          NextImg.Visible = False
        End If

        Dim i As Integer = 0
        Dim link As String = ""

        sb = New StringBuilder
        sb.Append("<div class=""lp"">")
        sb.Append("<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" summary=""問題列表資料表格"">")
        sb.Append("<tr><th scope=""col"">&nbsp;</th><th scope=""col"">標題</th><th scope=""col"">建立日期</th><th scope=""col"">討論數</th><th scope=""col"">評價</th>")
        sb.Append("<th scope=""col"">瀏覽數</th><th scope=""col"">討論關閉</th>")

        If ActivityFlag Then sb.Append("<th scope=""col"">活動問題</th>")

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
		  
		  
		  
          sb.Append("<td>" & link & "</td>")
          sb.Append("<td>[" & Date.Parse(myDataRow("xPostDate")).ToShortDateString & "]</td>")
          'sb.Append("<td>" & myDataRow("DiscussCount") & "</td>")
		  sb.Append("<td>" & totalDiscuss.ToString & "</td>")

          sb.Append("<td><div class=""staricon"">")
          sb.Append(ConcatStarString(Integer.Parse(myDataRow("GradeCount")), Integer.Parse(myDataRow("GradePersonCount"))))
          sb.Append("</div></td>")
          sb.Append("<td>" & myDataRow("BrowseCount") & "</td>")

          If myDataRow("xNewWindow") = "Y" Then
            sb.Append("<td>是</td>")
          Else
            sb.Append("<td>&nbsp;</td>")
          End If

          'If myDataRow("HavePros") = "Y" Then
          ' sb.Append("<td>V</td>")
          ' Else
          ' sb.Append("<td>&nbsp;</td>")
          ' End If

          If ActivityFlag Then
            If myDataRow("vGroup") = "A" And myDataRow("xNewWindow")="N" Then
                            sb.Append("<td><img src=""../xslGip/style3/images/good2.gif"" alt=""本問題為知識問答你我他活動問題"" /></td>")
            Else
              sb.Append("<td>&nbsp;</td>")
            End If
          End If

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

    Function ConcatStarString(ByVal GradeCount As Integer, ByVal PersonCount As Integer) As String

        Dim dresult As Double = 0

        If PersonCount = 0 Then
            dresult = GradeCount / 1
        Else
            dresult = GradeCount / PersonCount
        End If

        Dim str As String = ""
        If dresult = 0 Then
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 0 And dresult < 0.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 0.5 And dresult <= 1 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 1 And dresult < 1.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 1.5 And dresult <= 2 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 2 And dresult < 2.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 2.5 And dresult <= 3 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 3 And dresult < 3.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult >= 3.5 And dresult <= 4 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_empty_11x11.gif"" align=""top"">"
        ElseIf dresult > 4 And dresult < 4.5 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_half_11x11.gif"" align=""top"">"
        ElseIf dresult >= 4.5 And dresult <= 5.0 Then
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
            str &= "<img class=""rating"" src=""images/icn_star_full_11x11.gif"" align=""top"">"
        End If

        Return str

    End Function

End Class
