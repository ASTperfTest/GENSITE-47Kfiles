<%@ Page Language="vb" AutoEventWireup="false" Codebehind="a_publish_list.aspx.vb" Inherits="member.a_publish_list" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>a_publish_list</title>
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<!--<link href="../../css/intra.css" rel="stylesheet" type="text/css">-->
		<style type="text/css">BODY { MARGIN: 15px }
	.tt1 { FONT-WEIGHT: bold; FONT-SIZE: 16px; PADDING-BOTTOM: 10px; COLOR: #006699; FONT-FAMILY: "新細明體"; LETTER-SPACING: 0.1em; em: }
	.hr2 { BORDER-RIGHT: #aaaaaa dotted; BORDER-TOP: #aaaaaa dotted; BORDER-LEFT: #aaaaaa dotted; BORDER-BOTTOM: #aaaaaa dotted; HEIGHT: 3px }
	.Table1 { BORDER-RIGHT: #bbbbbb 2px solid; BORDER-TOP: #bbbbbb 2px solid; PADDING-BOTTOM: 10px; BORDER-LEFT: #bbbbbb 2px solid; WIDTH: 100%; BORDER-BOTTOM: #bbbbbb 2px solid; BACKGROUND-COLOR: #f4f4ea }
	.Table2 { BORDER-RIGHT: #999999 1px dotted; BORDER-TOP: #999999 1px dotted; MARGIN: 10px 0px 0px; BORDER-LEFT: #999999 1px dotted; WIDTH: 100%; BORDER-BOTTOM: #999999 1px dotted; BACKGROUND-COLOR: #f6f6f6 }
	.Item { PADDING-RIGHT: 0px; PADDING-LEFT: 10px; FONT-SIZE: 12px; PADDING-BOTTOM: 0px; VERTICAL-ALIGN: top; COLOR: #993333; LINE-HEIGHT: 150%; PADDING-TOP: 5px; FONT-FAMILY: "新細明體"; WHITE-SPACE: nowrap; TEXT-ALIGN: right }
	.Cont { PADDING-RIGHT: 10px; PADDING-LEFT: 0px; FONT-SIZE: 12px; PADDING-BOTTOM: 0px; VERTICAL-ALIGN: top; COLOR: #333333; LINE-HEIGHT: 150%; PADDING-TOP: 5px; FONT-FAMILY: "新細明體"; TEXT-ALIGN: left }
	.FormTd { PADDING-RIGHT: 10px; PADDING-LEFT: 0px; FONT-SIZE: 12px; PADDING-BOTTOM: 3px; VERTICAL-ALIGN: top; COLOR: #333333; PADDING-TOP: 3px; TEXT-ALIGN: left }
	.FormTx { FONT-SIZE: 12px; COLOR: #000000; LINE-HEIGHT: 140%; FONT-FAMILY: "新細明體" }
	.FormTx2 { FONT-SIZE: 12px; COLOR: #333333; LINE-HEIGHT: 140%; FONT-FAMILY: "新細明體"; BACKGROUND-COLOR: #eeeeee }
	.SubmitTd { PADDING-RIGHT: 10px; PADDING-LEFT: 10px; PADDING-BOTTOM: 10px; HEIGHT: 50px; TEXT-ALIGN: center }
	.ButtonTx { FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: "新細明體"; TEXT-DECORATION: none }
	.sep1 { BACKGROUND-COLOR: #cccccc }
	.Step { BORDER-RIGHT: 0px; PADDING-RIGHT: 0px; BORDER-TOP: 0px; MARGIN-TOP: 10px; PADDING-LEFT: 0px; FONT-SIZE: 13px; MARGIN-BOTTOM: 0px; PADDING-BOTTOM: 0px; MARGIN-LEFT: 10px; BORDER-LEFT: 0px; COLOR: #888888; LINE-HEIGHT: 110%; PADDING-TOP: 0px; BORDER-BOTTOM: 0px; FONT-FAMILY: "新細明體"; LETTER-SPACING: 0.1em }
	.StepNow { BORDER-RIGHT: #ffcccc 2px solid; PADDING-RIGHT: 5px; BORDER-TOP: #ffcccc 2px solid; PADDING-LEFT: 5px; FONT-SIZE: 15px; PADDING-BOTTOM: 0px; MARGIN: 0px; BORDER-LEFT: #ffcccc 2px solid; COLOR: #cc0033; PADDING-TOP: 4px; BORDER-BOTTOM: #ffcccc 2px solid }
	.tt2 { PADDING-RIGHT: 10px; PADDING-LEFT: 10px; FONT-SIZE: 15px; COLOR: #655316; PADDING-TOP: 10px; FONT-FAMILY: "新細明體"; LETTER-SPACING: 0.1em }
	.ListTable { BORDER-RIGHT: medium none; BORDER-TOP: medium none; FONT-SIZE: 12px; BORDER-LEFT: medium none; WIDTH: 100%; COLOR: #333333; LINE-HEIGHT: 150%; BORDER-BOTTOM: medium none }
	.ListHead { FONT-SIZE: 13px; COLOR: #005566; BACKGROUND-COLOR: #eeeeee; TEXT-ALIGN: center }
	.ListTitle { FONT-SIZE: 13px; COLOR: #003399 }
	.ListDate { VERTICAL-ALIGN: top; COLOR: #006666; FONT-FAMILY: "Arial", "Helvetica", "sans-serif"; TEXT-ALIGN: center }
	.ListNum { PADDING-RIGHT: 5px; FONT-WEIGHT: bold; VERTICAL-ALIGN: top; FONT-FAMILY: "Arial", "Helvetica", "sans-serif"; TEXT-ALIGN: center }
	.ListCenter { VERTICAL-ALIGN: top; TEXT-ALIGN: center }
	.ButNew { BORDER-RIGHT: medium none; BORDER-TOP: medium none; FONT-SIZE: 12px; MARGIN-LEFT: 5px; VERTICAL-ALIGN: middle; BORDER-LEFT: medium none; COLOR: #335599; BORDER-BOTTOM: medium none; TEXT-DECORATION: underline }
	A:link { COLOR: #003399 }
	A:hover { TEXT-DECORATION: none }
	.hr3 { BORDER-RIGHT: #336666 dotted; PADDING-RIGHT: 0px; BORDER-TOP: #336666 dotted; PADDING-LEFT: 0px; PADDING-BOTTOM: 0px; MARGIN: 0px; BORDER-LEFT: #336666 dotted; PADDING-TOP: 0px; BORDER-BOTTOM: #336666 dotted; HEIGHT: 1px }
	.Table3 { BORDER-RIGHT: #cccc99 1px solid; BORDER-TOP: #cccc99 1px solid; FONT-SIZE: 12px; MARGIN: 10px 0px; BORDER-LEFT: #cccc99 1px solid; BORDER-BOTTOM: #cccc99 1px solid; TEXT-ALIGN: center }
	.sep2 { BACKGROUND-COLOR: #cccc99 }
	.list_cont:link { FONT-SIZE: 12px; COLOR: #2d7f80; TEXT-DECORATION: none }
	.list_cont:hover { FONT-SIZE: 12px; COLOR: #993300; TEXT-DECORATION: none }
	.hr1 { BORDER-RIGHT: #999999 solid; BORDER-TOP: #999999 solid; MARGIN-BOTTOM: 5px; BORDER-LEFT: #999999 solid; BORDER-BOTTOM: #999999 solid; HEIGHT: 1px }
	.tt3 { FONT-SIZE: 15px; COLOR: #990000; FONT-FAMILY: "新細明體"; LETTER-SPACING: 0.1em }
	.green_cont { FONT-SIZE: 12px; COLOR: #2d7f80; TEXT-DECORATION: none }
		</style>
		<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
		</script>
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
			<table cellSpacing="0" cellPadding="0" width="100%" summary="系統名稱" border="0">
				<tr>
					<td class="tt1">學者會員審核
						<hr class="hr2">
					</td>
				</tr>
			</table>
			<table class="Table2" cellSpacing="5" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="Cont">
						<asp:DropDownList id="ddl_verify" runat="server" AutoPostBack="True">
							<asp:ListItem Value="All">全部</asp:ListItem>
							<asp:ListItem Value="N" Selected="True">未審核</asp:ListItem>
							<asp:ListItem Value="Y">已審核</asp:ListItem>
							<asp:ListItem Value="X">審核不通過</asp:ListItem>
						</asp:DropDownList>&nbsp;&nbsp;&nbsp;
						<asp:Button id="ButtonSelectAll" runat="server" Text="全選"></asp:Button>&nbsp;&nbsp;&nbsp;
						<asp:Button id="ButtonSelectNone" runat="server" Text="全不選"></asp:Button>&nbsp;&nbsp;&nbsp;
						<asp:Button id="Button2" runat="server" Text="同意"></asp:Button>&nbsp;&nbsp;&nbsp;
						<asp:Button id="Button3" runat="server" Text="不同意"></asp:Button>&nbsp;&nbsp;<asp:imagebutton id="ButtonAdd" runat="server" ImageUrl="../../images/icon_add.gif" AlternateText="新增"
							Visible="False"></asp:imagebutton></td>
				</tr>
			</table>
			<table class="Cont" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td>
						<TABLE class="Cont" id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD height="40"><asp:label id="LabelPage" runat="server"></asp:label><asp:label id="LabelCount" runat="server"></asp:label><asp:dropdownlist id="DropDownListSelectedPage" runat="server" AutoPostBack="True"></asp:dropdownlist><asp:label id="LabelUnitCount" runat="server"></asp:label></TD>
								<td align="right">&nbsp;&nbsp;&nbsp;<asp:image id="Image1" runat="server" ImageUrl="../../images/icon_new.gif" Visible="False"></asp:image>
									<asp:hyperlink id="HyperLink3" runat="server" Visible="False" NavigateUrl="a_publish_new.aspx">新增出版品</asp:hyperlink></td>
							</TR>
						</TABLE>
						<asp:Label id="LabelErr" runat="server" ForeColor="Red"></asp:Label>
					</td>
				</tr>
			</table>
			<asp:datagrid id="DataGrid1" runat="server" Width="100%" GridLines="None" Font-Size="12px" BorderColor="#DEDFDE"
				BorderStyle="None" BorderWidth="1px" BackColor="White" CellPadding="1" ForeColor="Black" AutoGenerateColumns="False"
				AllowPaging="True">
				<FooterStyle BackColor="#CCCC99"></FooterStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#CE5D5A"></SelectedItemStyle>
				<AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
				<ItemStyle BackColor="White"></ItemStyle>
				<HeaderStyle Height="30px" ForeColor="#005566" BackColor="#EEEEEE"></HeaderStyle>
				<Columns>
					<asp:TemplateColumn>
						<HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
						<ItemTemplate>
							<asp:CheckBox id="CheckBox1" runat="server"></asp:CheckBox>
						</ItemTemplate>
					</asp:TemplateColumn>
					<asp:BoundColumn Visible="False" DataField="account" SortExpression="account" HeaderText="登入帳號"></asp:BoundColumn>
					<asp:HyperLinkColumn DataNavigateUrlField="account" DataNavigateUrlFormatString="a_publish_modify.aspx?account={0}"
						DataTextField="account" SortExpression="account" HeaderText="登入帳號">
						<HeaderStyle Width="10%"></HeaderStyle>
					</asp:HyperLinkColumn>
					<asp:BoundColumn DataField="realname" SortExpression="realname" HeaderText="真實姓名">
						<HeaderStyle Width="10%"></HeaderStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="id" SortExpression="id" HeaderText="身份證字號">
						<HeaderStyle Width="10%"></HeaderStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="phone" SortExpression="phone" HeaderText="電話">
						<HeaderStyle Width="10%"></HeaderStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="fax" SortExpression="fax" HeaderText="傳真">
						<HeaderStyle Width="10%"></HeaderStyle>
					</asp:BoundColumn>
					<asp:BoundColumn DataField="email" SortExpression="email" HeaderText="電子郵件地址"></asp:BoundColumn>
					<asp:BoundColumn DataField="申請日期" HeaderText="申請日期">
						<HeaderStyle Width="10%"></HeaderStyle>
					</asp:BoundColumn>
					<asp:TemplateColumn Visible="False" HeaderText="修改">
						<HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
						<ItemTemplate>
							<asp:HyperLink id=HyperLink1 runat="server" ImageUrl="../../images/icon_edit.gif" NavigateUrl='<%# DataBinder.Eval(Container, "DataItem.ID", "a_publish_modify.aspx?id={0}") %>' Text="">
							</asp:HyperLink>
						</ItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn Visible="False" HeaderText="刪除">
						<HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center"></ItemStyle>
						<ItemTemplate>
							<asp:ImageButton id="ImageButtonDelete" runat="server" ImageUrl="../../images/icon_delete.gif" CommandName="delete"></asp:ImageButton>
						</ItemTemplate>
					</asp:TemplateColumn>
				</Columns>
				<PagerStyle Visible="False" HorizontalAlign="Center" ForeColor="Black" BackColor="White" Mode="NumericPages"></PagerStyle>
			</asp:datagrid></form>
	</body>
</HTML>
