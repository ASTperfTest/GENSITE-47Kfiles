Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Class SubjectList
  Inherits System.Web.UI.Page

  Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ODBCDSN").ConnectionString
  Dim sqlString As String = ""
    
  Dim myReader As SqlDataReader
  Dim myTable As DataTable
  Dim myDataRow As DataRow

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        getPageData(False)
        Me.t1.Attributes.Add("onkeypress", "if( event.keyCode == 13 ) {" & Me.ClientScript.GetPostBackEventReference(Me.btnSubmit, "") & "}")
    End Sub

    Protected Sub btnSubmit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSubmit.Click
        getPageData(True)
    End Sub

    Private Sub getPageData(ByVal IsSearch As Boolean)
        Dim SubjectId As String = ""
        Dim sb As New StringBuilder()

        If Request("SubjectId") = "" Then
            SubjectId = "all"
        Else
            SubjectId = Request("SubjectId")
        End If

        Dim PageSize As Integer = 10
        Dim PageNumber As Integer = 0

        If Not Page.IsPostBack Then
            Try
                If Request("PageSize") = "" Then
                    PageSize = 10
                Else
                    PageSize = CInt(Request("PageSize"))
                End If
                If Request("PageNumber") = "" Then
                    PageNumber = 1
                Else
                    PageNumber = CInt(Request("PageNumber"))
                End If
            Catch ex As Exception
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            End Try
        Else
            If Request.Form("__EVENTTARGET") <> Nothing And Request.Form("__EVENTTARGET").Replace("$", "_") = PageSizeDDL_.ClientID Then
                PageSizeDDL.SelectedValue = PageSizeDDL_.SelectedValue
            End If
            PageSize = PageSizeDDL.SelectedValue

            If Request.Form("__EVENTTARGET") <> Nothing And Request.Form("__EVENTTARGET").Replace("$", "_") = PageNumberDDL_.ClientID Then
                PageNumberDDL.SelectedValue = PageNumberDDL_.SelectedValue
            End If
            PageNumber = PageNumberDDL.SelectedValue

        End If

        Dim total As Integer = 0
        Dim pageCount As Integer = 0
        Dim position As Integer = 1


        Using myConnection As New SqlConnection(ConnString)

            Try
                getSqlString(SubjectId, IsSearch)

                'Response.Write(sqlString)

                myConnection.Open()
                Dim myCommand As SqlCommand = New SqlCommand(sqlString, myConnection)
                If SubjectId <> "all" Then
                    myCommand.Parameters.Add(New SqlParameter("@subjectid", SubjectId))
                End If
                If IsSearch Then
                    myCommand.Parameters.Add(New SqlParameter("@CtRootName", "%" + t1.Text.Trim() + "%"))
                End If
                myReader = myCommand.ExecuteReader()

                myTable = New DataTable()
                myTable.Load(myReader)

                total = myTable.Rows().Count()
                If total <> 0 Then
                    pageCount = Int(total / PageSize)
                    '判斷有沒有多 有多的才多加一頁
                    If (total Mod PageSize) > 0 Then
                        pageCount = pageCount + 1
                    End If

                    If pageCount < PageNumber Then
                        PageNumber = pageCount
                    End If

                    PageNumberText.Text = PageNumber.ToString()
                    PageNumberText_.Text = PageNumber.ToString()
                    TotalPageText.Text = pageCount.ToString()
                    TotalPageText_.Text = pageCount.ToString()
                    TotalRecordText.Text = total.ToString()
                    TotalRecordText_.Text = total.ToString()
                    If PageSize = 10 Then
                        PageSizeDDL.SelectedIndex = 0
                        PageSizeDDL_.SelectedIndex = 0
                    ElseIf PageSize = 20 Then
                        PageSizeDDL.SelectedIndex = 1
                        PageSizeDDL_.SelectedIndex = 1
                    ElseIf PageSize = 30 Then
                        PageSizeDDL.SelectedIndex = 2
                        PageSizeDDL_.SelectedIndex = 2
                    ElseIf PageSize = 50 Then
                        PageSizeDDL.SelectedIndex = 3
                        PageSizeDDL_.SelectedIndex = 3
                    End If
                    Dim item As ListItem
                    Dim j As Integer = 0
                    PageNumberDDL.Items.Clear()
                    PageNumberDDL_.Items.Clear()
                    For j = 0 To pageCount - 1
                        item = New ListItem
                        item.Value = j + 1
                        item.Text = j + 1
                        If PageNumber = (j + 1) Then
                            item.Selected = True
                        End If
                        PageNumberDDL.Items.Insert(j, item)
                        PageNumberDDL_.Items.Insert(j, item)
                        item = Nothing
                    Next
                    position = PageSize * (PageNumber - 1)
                    myDataRow = myTable.Rows.Item(position)

                    If PageNumber > 1 Then
                        PreviousText.Visible = True
                        PreviousText_.Visible = True
                        PreviousImg.Visible = True
                        PreviousImg_.Visible = True
                        'PreviousLink.NavigateUrl = "SubjectList.aspx?SubjectId=" & SubjectId & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                        PreviousLink.NavigateUrl = "javascript:document.getElementById('" & PageNumberDDL.ClientID & "').selectedIndex--;document.forms['aspnetForm'].submit()"
                        PreviousLink_.NavigateUrl = "javascript:document.getElementById('" & PageNumberDDL.ClientID & "').selectedIndex--;document.forms['aspnetForm'].submit()" '2011.07.08 Grace
                    Else
                        PreviousText.Visible = False
                        PreviousText_.Visible = False
                        PreviousImg.Visible = False
                        PreviousImg_.Visible = False
                    End If
                    If PageNumberDDL.SelectedValue < pageCount And PageNumberDDL_.SelectedValue < pageCount Then
                        NextText.Visible = True
                        NextText_.Visible = True
                        NextImg.Visible = True
                        NextImg_.Visible = True
                        NextLink.NavigateUrl = "SubjectList.aspx?SubjectId=" & SubjectId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                        NextLink_.NavigateUrl = "SubjectList.aspx?SubjectId=" & SubjectId & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize '2011.07.08 Grace
                        NextLink.NavigateUrl = "javascript:document.getElementById('" & PageNumberDDL.ClientID & "').selectedIndex++;document.forms['aspnetForm'].submit()"
                        NextLink_.NavigateUrl = "javascript:document.getElementById('" & PageNumberDDL.ClientID & "').selectedIndex++;document.forms['aspnetForm'].submit()" '2011.07.08 Grace
                    Else
                        NextText.Visible = False
                        NextText_.Visible = False
                        NextImg.Visible = False
                        NextImg_.Visible = False
                    End If

                    Dim i As Integer = 0
                    Dim link As String = ""
                    'Dim CTime As Array
                    Dim atime As Date
                    Dim utime As Date

                    sb.Append("<div class=""subjectlist""><ul>")

                    For i = 0 To PageSize - 1

                        link = ""
                        sb.Append("<li>")

                        atime = CDate(myDataRow.Item("create_time"))
                        utime = CDate(myDataRow.Item("update_time"))

                        If myDataRow.Item("url_link").ToString() <> "" Then
                            link = "<a href=""" & myDataRow.Item("url_link").ToString() & """ target=""_blank"">"
                            link &= "<img src=""/public/" & myDataRow.Item("pic").ToString() & """ alt=""" & myDataRow.Item("CtRootID").ToString() & """ border=""0"" /></a>"
                            link &= "<a href=""" & myDataRow.Item("url_link").ToString() & """ target=""_blank"">" & myDataRow.Item("CtRootName").ToString() & "</a>"
                        Else
                            link = "<a href=""/subject/mp.asp?mp=" & myDataRow.Item("CtRootID").ToString() & """ target=""_blank"">"
                            link &= "<img src=""/public/" & myDataRow.Item("pic").ToString() & """ alt=""" & myDataRow.Item("CtRootID").ToString() & """ border=""0"" /></a>"
                            link &= "<a href=""/subject/mp.asp?mp=" & myDataRow.Item("CtRootID").ToString() & """ target=""_blank"">" & myDataRow.Item("CtRootName").ToString() & "</a>"
                        End If
                        'link  &="<a href=""/subject/mp.asp?mp=" & myDataRow.Item("CtRootID").ToString() & """ target=""_blank"">" & myDataRow.Item("CtRootName").ToString() & "</a>"
                        'link &= CTime(2) & "-" & CTime(0) & "-" & CTime(1)
                        link &= "建罝日期："
                        link &= atime.Year & "-" & atime.Month & "-" & atime.Day
                        link &= "<font color=white>"
                        link &= "&nbsp;更新時間："
                        link &= utime.Year & "-" & utime.Month & "-" & utime.Day
                        link &= "&nbsp;瀏覽次數："
                        If Not DBNull.Value.Equals(myDataRow.Item("hits")) Then
                            link &= myDataRow.Item("hits")
                        Else
                            link &= "0"
                        End If
                        link &= "</font>"
                        link &= "<span>" & myDataRow.Item("abstract").ToString() & "</span>"
                        sb.Append(link)
                        sb.Append("</li><hr/>")

                        position += 1

                        If myTable.Rows.Count <= position Then
                            Exit For
                        End If
                        myDataRow = myTable.Rows.Item(position)
                    Next

                    sb.Append("</ul></div>")

                    TableText.Text = sb.ToString()
                Else
                    PageNumberText.Text = ""
                    PageNumberText_.Text = ""
                    TotalPageText.Text = ""
                    TotalPageText_.Text = ""
                    TotalRecordText.Text = "0"
                    TotalRecordText_.Text = "0"
                    NextText.Visible = False
                    NextText_.Visible = False
                    NextImg.Visible = False
                    NextImg_.Visible = False
                    PageNumberDDL.Items.Clear()
                    PageNumberDDL_.Items.Clear()
                    PageNumberDDL.Items.Insert(0, 1)
                    PageNumberDDL_.Items.Insert(0, 1)
                    TableText.Text = ""
                End If
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
                myTable.Dispose()
            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try

            '---tab list---
            sb = Nothing
            sb = New StringBuilder
            Try
                sqlString = "SELECT classid, classname FROM Type WHERE (datalevel = 1) "
                myConnection.Open()
                Dim myCommand As SqlCommand = New SqlCommand(sqlString, myConnection)
                myReader = myCommand.ExecuteReader

                sb.AppendLine("<ul class=""subjectcata"">")

                If SubjectId = "all" Then
                    sb.AppendLine("<li class=""active"">")
                Else
                    sb.AppendLine("<li>")
                End If
                sb.AppendLine("<a href=""SubjectList.aspx?SubjectId=all"" alt=""全部"">全部</a>")
                sb.AppendLine("</li>")

                While myReader.Read

                    If myReader.Item("classid").ToString() = SubjectId Then
                        sb.AppendLine("<li class=""active"">")
                    Else
                        sb.AppendLine("<li>")
                    End If
                    sb.AppendLine("<a href=""SubjectList.aspx?SubjectId=" & myReader.Item("classid") & """ alt=""" & myReader.Item("classname") & """>" & myReader.Item("classname") & "</a>")

                    sb.AppendLine("</li>")
                End While
                sb.AppendLine("</ul>")

                TabText.Text = sb.ToString()
                If Not myReader.IsClosed Then
                    myReader.Close()
                End If
            Catch ex As Exception
                If Request("debug") = "true" Then
                    Response.Write(ex.ToString())
                    Response.End()
                End If
                Response.Redirect("/mp.asp?mp=1")
                Response.End()
            Finally
                If myConnection.State = ConnectionState.Open Then
                    myConnection.Close()
                End If
            End Try
        End Using
    End Sub

    Private Function getSqlString(ByVal SubjectId As String, ByVal IsSearch As Boolean) As String
        sqlString = "SELECT "
        sqlString &= "      CatTreeRoot.CtRootID "
        sqlString &= "     , CatTreeRoot.CtRootName "
        sqlString &= "     , NodeInfo.create_time "
        sqlString &= "     , NodeInfo.pic "
        sqlString &= "     , Type.classid "
        sqlString &= "     , NodeInfo.abstract "
        sqlString &= "     , NodeInfo.url_link  "
        'If SortDDL.SelectedValue = 1 Then
        'sqlString &= "     ,( "
        'sqlString &= "         SELECT  "
        'sqlString &= "             SUM(DailyClick.dailyClick) as sum2  "
        'sqlString &= "         FROM CatTreeNode  "
        'sqlString &= "         INNER JOIN CuDTGeneric ON CatTreeNode.CtUnitID = CuDTGeneric.iCTUnit  "
        'sqlString &= "         LEFT JOIN DailyClick ON CuDTGeneric.iCUItem = DailyClick.iCUItem  "
        'sqlString &= "         WHERE CatTreeNode.CtRootID = CatTreeRoot.CtRootID "
        'sqlString &= "     ) hits "
        'End If
        sqlString &= "     ,oInfo.hits "
        sqlString &= "     ,oInfo.[文章篇數] "
        sqlString &= "     ,oInfo.update_time "
        sqlString &= " FROM NodeInfo  "
        sqlString &= " INNER JOIN CatTreeRoot ON CatTreeRoot.CtRootID = NodeInfo.CtrootID  "
        sqlString &= " INNER JOIN [Type] ON NodeInfo.type1 = Type.classname  "
        sqlString &= " left join  "
        sqlString &= " ( "
        sqlString &= "     SELECT  "
        sqlString &= "           M2.CtRootID "
        sqlString &= "         , COUNT(*) as [文章篇數]   "
        sqlString &= "         , sum(CuDTGeneric.ClickCount) as hits  "
        sqlString &= "         , ( "
        sqlString &= "             select max(CuDTGeneric.dEditDate)  "
        sqlString &= "             from CuDTGeneric  "
        sqlString &= "             where CuDTGeneric.iCTUnit in (select CtUnitID FROM CatTreeNode WHERE CatTreeNode.CtRootID = M2.CtRootID ) "
        sqlString &= "           ) as update_time "
        sqlString &= "     FROM CatTreeNode M2 INNER JOIN CuDTGeneric ON M2.CtUnitID = CuDTGeneric.iCTUnit  "
        sqlString &= "     group by M2.CtRootID "
        sqlString &= " ) oInfo on oInfo.CtRootID = CatTreeRoot.CtrootID  "
        sqlString &= "   "
        sqlString &= " WHERE (CatTreeRoot.inUse = 'Y')  "
        If SubjectId <> "all" Then
            sqlString &= " AND (Type.classid = @subjectid) "
        End If
        If IsSearch Then
            sqlString &= " AND (CatTreeRoot.CtRootName like @CtRootName ) "
        End If

        sqlString &= " order by  "
        If SortDDL.SelectedValue = 0 Then
            sqlString &= "     NEWID(),"
        ElseIf SortDDL.SelectedValue = 3 Then
            sqlString &= "     oInfo.update_time desc,"
        ElseIf SortDDL.SelectedValue = 1 Then
            sqlString &= "     hits desc,"
        ElseIf SortDDL.SelectedValue = 2 Then
            sqlString &= "     oInfo.[文章篇數] desc,"
        End If

        sqlString &= "     (Case When NodeInfo.order_num is null Then 1 Else 0 End) "
        sqlString &= "     , order_num DESC "
        sqlString &= "     , CatTreeRoot.CtrootID  "
        sqlString &= " "

        Return sqlString
    End Function


    Protected Sub SortDDL_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles SortDDL.SelectedIndexChanged
        t1.Text = ""
    End Sub
    Protected Sub PageSizeDDL_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles PageSizeDDL.SelectedIndexChanged
        t1.Text = ""
    End Sub
    Protected Sub PageSizeDDL__SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles PageSizeDDL_.SelectedIndexChanged
        t1.Text = ""
    End Sub '2011.07.08 Grace

End Class
