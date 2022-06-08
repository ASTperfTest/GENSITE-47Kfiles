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
		<body>
		<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
			<tr>
				<td width="50%" class="FormName" align="left">–る计参璸&nbsp;<font size="2">–る计参璸</td>
				<td width="50%" class="FormLink" align="right">
				</td>
			</tr>
			<tr>
				<td width="100%" colspan="2">
					<hr noshade size="1" color="#000080">
				</td>
			</tr>
			
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:DataGrid id="DataGrid1" runat="server"
			 OnPageIndexChanged="Grid_Change"
			 BorderColor="Tan" BackColor="LightGoldenrodYellow" BorderWidth="1px" CellPadding="2"
				AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" Width="216px" GridLines="None"
				ForeColor="Black">
				<FooterStyle BackColor="Tan"></FooterStyle>
				<SelectedItemStyle ForeColor="GhostWhite" BackColor="DarkSlateBlue"></SelectedItemStyle>
				<AlternatingItemStyle BackColor="PaleGoldenrod"></AlternatingItemStyle>
				<HeaderStyle Font-Bold="True" BackColor="Tan"></HeaderStyle>
				<Columns>
					<asp:BoundColumn Visible="False" DataField="ID" ReadOnly="True" HeaderText="ID"></asp:BoundColumn>
					<asp:BoundColumn DataField="TmpTime" HeaderText="丁"></asp:BoundColumn>
					<asp:BoundColumn DataField="Person" HeaderText="计"></asp:BoundColumn>
				</Columns>
					<PagerStyle  HorizontalAlign="Center" ForeColor="DarkSlateBlue" 
					Position="TopAndBottom" BackColor="PaleGoldenrod"></PagerStyle>
			</asp:DataGrid>
		</form>
			<tr>
				<td width="100%" colspan="2" align="center">
				</td>
			</tr>
		</table>
	</body>
</HTML>
<script language="VB" runat="server">

Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) 
      With DataGrid1
            ' Enable paging.
            .AllowPaging = True
            ' Display 5 page numbers at a time.
            .PagerStyle.Mode = PagerMode.NumericPages
            .PagerStyle.PageButtonCount = 5
            .PageSize = 5
        End With
        
      If Not Page.IsPostBack Then
        Dim mySelectQuery As String
        Dim myConnection As New SqlConnection(ConfigurationSettings.AppSettings("ConnString"))

        myConnection.Open()
        Dim sqlstring As String = " SELECT  MonthPerson.ID, MonthPerson.TmpTime , MonthPerson.person FROM  MonthPerson "
      

        Dim myCommand As SqlDataAdapter = New SqlDataAdapter(sqlstring, myConnection)
        Dim ds2 As DataSet = New DataSet
        myCommand.Fill(ds2, "MonthPerson")

        DataGrid1.DataSource = ds2.Tables("MonthPerson")

        DataGrid1.DataBind()
       end if 
    End Sub
    
    Sub Grid_Change(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) 
     
            Dim mySelectQuery As String
            Dim myConnection As New SqlConnection(ConfigurationSettings.AppSettings("ConnString"))

            myConnection.Open()
            Dim sqlstring As String = " SELECT  MonthPerson.ID, MonthPerson.TmpTime , MonthPerson.person FROM  MonthPerson "
   

            Dim myCommand As SqlDataAdapter = New SqlDataAdapter(sqlstring, myConnection)
            Dim ds2 As DataSet = New DataSet
            myCommand.Fill(ds2, "MonthPerson")

            DataGrid1.DataSource = ds2.Tables("MonthPerson")

        DataGrid1.CurrentPageIndex = e.NewPageIndex
        DataGrid1.DataBind()
    End Sub

</script>

