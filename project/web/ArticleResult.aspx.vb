Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Configuration

Partial Public Class ArticleResult
    Inherits System.Web.UI.Page

    Dim ConnString As String = WebConfigurationManager.ConnectionStrings("ODBCDSN").ConnectionString
    Dim sqlString As String = ""
    Dim myConnection As SqlConnection
    Dim myCommand As SqlCommand
    Dim myReader As SqlDataReader
    Dim myTable As DataTable
    Dim myDataRow As DataRow

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim SubjectName As String = ""
        Dim Title As String = ""
        Dim Body As String = ""
        Dim sDate As String = ""
        Dim Type As String = ""
        Dim sb As New StringBuilder()

        SubjectName = Server.UrlDecode(Request("SubjectName"))
        Title = Server.UrlDecode(Request("Title"))
        Body = Server.UrlDecode(Request("Body"))
        sDate = Request("PDate")
        Type = Server.UrlDecode(Request("Type"))

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
            PageSize = PageSizeDDL.SelectedValue
            PageNumber = PageNumberDDL.SelectedValue
        End If

        Dim total As Integer = 0
        Dim pageCount As Integer = 0
        Dim position As Integer = 1

        myConnection = New SqlConnection(ConnString)

        Try
            getSqlString()
            'Response.Write(sqlString)
            myConnection.Open()
            myCommand = New SqlCommand(sqlString, myConnection)
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
                TotalPageText.Text = pageCount.ToString()
                TotalRecordText.Text = total.ToString()
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
                For j = 0 To pageCount - 1
                    item = New ListItem
                    item.Value = j + 1
                    item.Text = j + 1
                    If PageNumber = (j + 1) Then
                        item.Selected = True
                    End If
                    PageNumberDDL.Items.Insert(j, item)
                    item = Nothing
                Next
                position = PageSize * (PageNumber - 1)
                myDataRow = myTable.Rows.Item(position)

                If PageNumber > 1 Then
                    PreviousText.Visible = True
                    PreviousImg.Visible = True
                    PreviousLink.NavigateUrl = "ArticleResult.aspx?SubjectName=" & SubjectName & "&Title=" & Title & "&Body=" & Body & "&pDate=" & sDate & "&Type=" & Type & "&PageNumber=" & PageNumber - 1 & "&PageSize=" & PageSize
                Else
                    PreviousText.Visible = False
                    PreviousImg.Visible = False
                End If
                If PageNumberDDL.SelectedValue < pageCount Then
                    NextText.Visible = True
                    NextImg.Visible = True
                    NextLink.NavigateUrl = "ArticleResult.aspx?SubjectName=" & SubjectName & "&Title=" & Title & "&Body=" & Body & "&pDate=" & sDate & "&Type=" & Type & "&PageNumber=" & PageNumber + 1 & "&PageSize=" & PageSize
                Else
                    NextText.Visible = False
                    NextImg.Visible = False
                End If

                Dim i As Integer = 0
                Dim link As String = ""

                'sb.Append("<div class=""subjectlist""><ul>")
                sb.Append("<div class=""subjectlist""><table width='100%' border='0' cellpadding='0' cellspacing='1' style='background-color: #b2ebef'>")
                sb.Append("<tr style='background-color: #b2ebef'>")
                sb.Append("<td width='20%' height='30px' style='padding-left:10px;'><b><font color='#2c5858'>主題館</font></b></td>")
                sb.Append("<td style='padding-left:10px;'><b><font color='#2c5858'>文章標題</font></b></td>")
                sb.Append("<td width='20%' style='padding-left:10px;'><b><font color='#2c5858'>文章日期</font></b></td></tr>")
                For i = 0 To PageSize - 1
                    link = ""
                    'sb.Append("<li>")
                    sb.Append("<tr style='background-color: #FFFFFF' height='35px'>")
                    'If myDataRow.Item("url_link").ToString() <> "" Then
                    '    link &= "<a href=""" & myDataRow.Item("url_link").ToString() & """ target=""_blank"">" & myDataRow.Item("CtRootName").ToString() & "</a>"
                    'Else
                    link &= "<td style='padding-left:10px;'>"
                    link &= "<a href=""/subject/mp.asp?mp=" & myDataRow.Item("CtRootID").ToString() & """ target=""_blank"">" & myDataRow.Item("CtRootName").ToString() & "</a>"
                    link &= "</td>"
                    'End If
                    'link &= "&nbsp&nbsp&nbsp"
                    link &= "<td style='padding-left:10px;'>"
                    If Not DBNull.Value.Equals(myDataRow.Item("xURL")) Then
                        link &= "<a href=" & myDataRow.Item("xURL").ToString() & " target=_blank>" & myDataRow.Item("sTitle").ToString() & "</a>"
                    Else
                        link &= "<a href=""/subject/ct.asp?xItem=" & myDataRow.Item("iCUItem").ToString() & "&ctNode=" & myDataRow.Item("CtNodeId").ToString() & "&mp=" & myDataRow.Item("CtrootId").ToString() & "&kpi=0"" target=""_blank"">" & myDataRow.Item("sTitle").ToString() & "</a>"
                    End If
                    'link &= "&nbsp&nbsp&nbsp"
                    link &= "</td>"
                    link &= "<td style='padding-left:10px;'>"
                    link &= Convert.ToDateTime(myDataRow.Item("xPostDate")).ToString("yyyy/MM/dd")
                    link &= "</td>"
                    sb.Append(link)
                    'sb.Append("</li><hr/>")
                    sb.Append("</tr>")
                    position += 1

                    If myTable.Rows.Count <= position Then
                        Exit For
                    End If
                    myDataRow = myTable.Rows.Item(position)
                Next
                'sb.Append("</ul></div>")
                sb.Append("</table>")

                TableText.Text = sb.ToString()
            Else
                TotalRecordText.Text = "0"
                PageSizeDDL.Items.Clear()
                PageNumberDDL.Items.Clear()
            End If
            If Not myReader.IsClosed Then
                myReader.Close()
            End If
            myTable.Dispose()
            myCommand.Dispose()
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
    End Sub

    Private Function getSqlString() As String
        sqlString = "select "
        sqlString &= "	 CatTreeRoot.CtrootId "
        sqlString &= "	,CatTreeRoot.CtRootName "
        sqlString &= "	,CatTreeNode.CtNodeId "
        sqlString &= "	,NodeInfo.Type1 "
        sqlString &= "	,CuDTGeneric.*  "
        sqlString &= "from CuDTGeneric   "
        sqlString &= "inner join CatTreeNode on CatTreeNode.CtUnitID = CuDTGeneric.icTUnit  "
        sqlString &= "inner join CatTreeRoot on CatTreeNode.CtRootID = CatTreeRoot.CtrootID "
        sqlString &= "INNER JOIN NodeInfo    ON CatTreeNode.CtRootID = NodeInfo.CtrootID  "
        sqlString &= " "
        sqlString &= "where "
        sqlString &= "iCTUnit in  "
        sqlString &= "( "
        sqlString &= "	select  "
        sqlString &= "		CtUnitId  "
        sqlString &= "	from CatTreeNode  "
        sqlString &= "	where ctRootId in "
        sqlString &= "	( "
        sqlString &= "		SELECT  "
        sqlString &= "			CatTreeRoot.ctRootId "
        sqlString &= "		FROM NodeInfo  "
        sqlString &= "		INNER JOIN CatTreeRoot ON CatTreeRoot.CtRootID = NodeInfo.CtrootID  "
        sqlString &= "		WHERE (CatTreeRoot.inUse = 'Y')    "
        If Request("SubjectName") <> "" Then
            sqlString &= "		and CatTreeRoot.CtRootName like '%" + Server.UrlDecode(Request("SubjectName")) + "%'"
        End If
        If Request("Type") <> "" Then
            sqlString &= "		and NodeInfo.Type1 = '" + Server.UrlDecode(Request("Type")) + "'"
        End If
        If Request("Title") <> "" Then
            sqlString &= "		and CuDTGeneric.sTitle like '%" + Server.UrlDecode(Request("Title")) + "%'"
        End If
        If Request("Body") <> "" Then
            sqlString &= "		and CuDTGeneric.xBody like '%" + Server.UrlDecode(Request("Body")) + "%'"
        End If
        If Request("PDate").Length = 16 Then
            sqlString &= "		and xPostDate >= '" + Request("PDate").Substring(0, 8) + " 00:00:00' and xPostDate <= '" + Request("PDate").Substring(8, 8) + " 23:59:59'"
        End If
        sqlString &= "	) "
        sqlString &= ") "
        sqlString &= "and fCTUPublic = 'Y' "
        sqlString &= "and CuDTGeneric.sTitle is not null "

        Return sqlString
    End Function

End Class