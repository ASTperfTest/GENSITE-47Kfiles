<%@ Page Language="VB" AutoEventWireup="True" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace = "System.Data.SqlClient" %>
<%@ Import Namespace = "System.Data" %>
<%@ Import Namespace = "System.Web" %>


<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">

	
		<link rel="stylesheet" href="/inc/setstyle.css">
			<title></title>
	</HEAD>
		<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
		<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
			<tr>
				<td width="50%" class="FormName" align="left">KM计参璸&nbsp;<font size="2">–る计参璸</td>
				<td width="50%" class="FormLink" align="right">
				</td>
			</tr>
			<tr>
				<td width="100%" colspan="2">
					<hr noshade size="1" color="#000080">
				</td>
			</tr>
			
			<tr>
			<td>
		
			<FONT face="穝灿砰">
			<asp:DropDownList id="DropDownList2" AutoPostBack="True"  
					runat="server" BackColor="White" Width="100px">
					<asp:ListItem Value="1" Selected="True">禣</asp:ListItem>
					<asp:ListItem Value="2">ネ玻</asp:ListItem>
					<asp:ListItem Value="3">厩</asp:ListItem>
					<asp:ListItem Value="4">场</asp:ListItem>
				</asp:DropDownList>
			<asp:dropdownlist id="DropDownList1" 
					runat="server" Width="100px" AutoPostBack="True" OnSelectedIndexChanged="dropdownlist_Change">
						  <asp:ListItem  Value=1 Selected=True>る</asp:ListItem>
                          <asp:ListItem Value=2></asp:ListItem>
				</asp:dropdownlist>
						
				<asp:datagrid id="DataGrid1" 
					runat="server" Width="352px" BorderColor="#CCCCCC" BorderWidth="1px" BackColor="White"
					CellPadding="3" BorderStyle="None" PageSize="12" AllowPaging="True" OnPageIndexChanged="Grid_Change"
					AutoGenerateColumns="False">
					<FooterStyle ForeColor="#000066" BackColor="White"></FooterStyle>
					<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#669999"></SelectedItemStyle>
					<ItemStyle ForeColor="#000066"></ItemStyle>
					<HeaderStyle Font-Bold="True" ForeColor="White" BackColor="#006699"></HeaderStyle>
					
				
					<Columns>
						<asp:BoundColumn DataField="TmpTime" HeaderText=""></asp:BoundColumn>
						<asp:BoundColumn DataField="MonthTmp" HeaderText="る"></asp:BoundColumn>
						<asp:BoundColumn DataField="Person" HeaderText="计"></asp:BoundColumn>
					</Columns>
					<PagerStyle HorizontalAlign="Left" ForeColor="#000066" BackColor="White" Mode="NumericPages"></PagerStyle>
				</asp:datagrid>			
				
		</FONT>
		</td>
		</tr>
		</table>
		</form>
	</body>
</HTML>
<script language="VB" runat="server">

 Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        '硂柑竚ㄏノ祘Α絏﹍て呼
       Dim mySelectQuery As String
        Dim myConnection As New SqlConnection(ConfigurationSettings.AppSettings("ConnString"))
        Dim sqlstring As String
        myConnection.Open()
        Try
            DataGrid1.CurrentPageIndex = 0

            If DropDownList1.SelectedIndex <> -1 Then

                If DropDownList1.SelectedValue = 2 Then '

                    DataGrid1.Columns.Item(1).Visible = False
                    sqlstring = " SELECT   YearTmp as tmpTime , SUM(Person) AS Person, YearTmp as Monthtmp FROM  yearStatics"
                    ' sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    Else
                        sqlstring += ""

                    End If
                    sqlstring += " GROUP BY  YearTmp ORDER BY  YearTmp DESC  "

                Else 'る
                    DataGrid1.Columns.Item(1).Visible = True
                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring = " SELECT   YearTmp as TmpTime, MonthTmp, Person FROM  yearStatics"
                    Else
                        sqlstring = "SELECT SUM(Person) AS Person, YearTmp as TmpTime, MonthTmp FROM  yearStatics"
                    End If

                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    Else
                        sqlstring += ""

                    End If

                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " "
                    Else
                        sqlstring += " GROUP BY  YearTmp, MonthTmp "
                    End If
                    sqlstring += " ORDER BY  YearTmp DESC, MonthTmp DESC "
                End If



                'Response.Write(sqlstring)
                Dim myCommand As SqlDataAdapter = New SqlDataAdapter(sqlstring, myConnection)
                Dim ds2 As DataSet = New DataSet
                myCommand.Fill(ds2, "MonthPerson")

                DataGrid1.DataSource = ds2.Tables("MonthPerson")
             
                DataGrid1.DataBind()

            End If
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try
        '  End If
    End Sub
    Sub Grid_Change(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs)

        Dim mySelectQuery As String
        Dim myConnection As New SqlConnection(ConfigurationSettings.AppSettings("ConnString"))
        Dim sqlstring As String
        myConnection.Open()
        Try
           If DropDownList1.SelectedIndex <> -1 Then

                If DropDownList1.SelectedValue = 2 Then '

                    DataGrid1.Columns.Item(1).Visible = False
                    sqlstring = " SELECT   YearTmp as tmpTime , SUM(Person) AS Person, YearTmp as Monthtmp FROM  yearStatics"
                    ' sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    Else
                        sqlstring += ""

                    End If
                    sqlstring += " GROUP BY  YearTmp ORDER BY  YearTmp DESC  "

                Else 'る
                    DataGrid1.Columns.Item(1).Visible = True
                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring = " SELECT   YearTmp as TmpTime, MonthTmp, Person FROM  yearStatics"
                    Else
                        sqlstring = "SELECT SUM(Person) AS Person, YearTmp as TmpTime, MonthTmp FROM  yearStatics"
                    End If

                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    Else
                        sqlstring += ""

                    End If

                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " "
                    Else
                        sqlstring += " GROUP BY  YearTmp, MonthTmp "
                    End If
                    sqlstring += " ORDER BY  YearTmp DESC, MonthTmp DESC "
                End If




                '      Response.Write(sqlstring)
                Dim myCommand As SqlDataAdapter = New SqlDataAdapter(sqlstring, myConnection)
                Dim ds2 As DataSet = New DataSet
                myCommand.Fill(ds2, "MonthPerson")

                DataGrid1.DataSource = ds2.Tables("MonthPerson")
                If DataGrid1.CurrentPageIndex >= 0 And DataGrid1.CurrentPageIndex < DataGrid1.PageCount Then
                    DataGrid1.CurrentPageIndex = e.NewPageIndex
                Else

                    DataGrid1.CurrentPageIndex = 0
                End If
                DataGrid1.DataBind()

            End If
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

    End Sub


    Private Sub dropdownlist_Change(ByVal sender As System.Object, ByVal e As System.EventArgs) 

        Dim mySelectQuery As String
        Dim myConnection As New SqlConnection(ConfigurationSettings.AppSettings("ConnString"))
        Dim sqlstring As String
        myConnection.Open()
        Try
            If DropDownList1.SelectedIndex <> -1 Then

                If DropDownList1.SelectedValue = 2 Then '

                    DataGrid1.Columns.Item(1).Visible = False
                    sqlstring = " SELECT   YearTmp as tmpTime , SUM(Person) AS Person, YearTmp as Monthtmp FROM  yearStatics"
                    ' sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    Else
                        sqlstring += ""

                    End If
                    sqlstring += " GROUP BY  YearTmp ORDER BY  YearTmp DESC  "

                Else 'る
                    DataGrid1.Columns.Item(1).Visible = True
                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring = " SELECT   YearTmp as TmpTime, MonthTmp, Person FROM  yearStatics"
                    Else
                        sqlstring = "SELECT SUM(Person) AS Person, YearTmp as TmpTime, MonthTmp FROM  yearStatics"
                    End If

                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " where  mp='" & DropDownList2.SelectedValue & "' "
                    Else
                        sqlstring += ""

                    End If

                    If DropDownList2.SelectedValue <> "4" Then
                        sqlstring += " "
                    Else
                        sqlstring += " GROUP BY  YearTmp, MonthTmp "
                    End If
                    sqlstring += " ORDER BY  YearTmp DESC, MonthTmp DESC "
                End If




                'Response.Write(sqlstring)
                Dim myCommand As SqlDataAdapter = New SqlDataAdapter(sqlstring, myConnection)
                Dim ds2 As DataSet = New DataSet
                myCommand.Fill(ds2, "MonthPerson")

                DataGrid1.DataSource = ds2.Tables("MonthPerson")

                DataGrid1.DataBind()
            End If
        Catch ex As Exception
            Response.Write(ex.Message)
        End Try

    End Sub
</script>

