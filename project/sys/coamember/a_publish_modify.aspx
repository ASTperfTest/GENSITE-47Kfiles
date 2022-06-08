<%@ Page Language="vb" AutoEventWireup="false" Codebehind="a_publish_modify.aspx.vb" Inherits="member.a_publish_modify" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>a_publish_modify</title>
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<meta http-equiv="Content-Type" content="text/html; charset=big5">
		<!--<LINK href="../../css/intra.css" type="text/css" rel="stylesheet">-->
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
		<script language="javascript" src="../../include/SS_popup.js"></script>
	</HEAD>
	<body>
		<form id="Form1" method="post" encType="multipart/form-data" runat="server">
			<table cellSpacing="0" cellPadding="0" width="100%" summary="系統名稱" border="0">
				<tr>
					<td class="tt1">會員詳細資料&nbsp;
						<hr class="hr2">
					</td>
				</tr>
			</table>
			<table cellSpacing="5" cellPadding="0" width="100%" border="0">
				<tr>
					<td></td>
				</tr>
			</table>
			<table class="Table1">
				<form>
					<TBODY>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">會員帳號：</td>
							<td>
								<asp:Label id="LabelAccount" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">真實姓名：</td>
							<td>
								<asp:Label id="LabelRealname" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">身份證字號：</td>
							<td>
								<asp:Label id="LabelIDN" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" vAlign="top" align="right">身份類別：</td>
							<td>
								<asp:Label id="LabelActor" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">所屬機關名稱：</td>
							<td>
								<asp:Label id="LabelMember_org" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" vAlign="top" align="right">所屬機關電話：</td>
							<td>
								<asp:Label id="LabelCom_tel" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" vAlign="top" align="right">職稱：</td>
							<td>
								<asp:Label id="LabelPtitle" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">出生日期：</td>
							<td>
								<asp:Label id="LabelBirthday" runat="server">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">性別：</td>
							<td>
								<asp:Label id="LabelSex" runat="server">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">地址：</td>
							<td>
								<asp:Label id="LabelHomeAddr" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px; HEIGHT: 31px" vAlign="top" align="right">電話：</td>
							<td>
								<asp:Label id="LabelPhone" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">手機號碼：</td>
							<td>
								<asp:Label id="LabelMobile" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<tr>
						<tr>
							<td class="Item" style="WIDTH: 99px" align="right">傳真：</td>
							<td>
								<asp:Label id="LabelFax" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
						<TR>
							<td class="Item" style="WIDTH: 99px" vAlign="top" align="right">電子郵件地址：</td>
							<td>
								<asp:Label id="LabelEmail" runat="server" Font-Size="10">Label</asp:Label></td>
						</TR>
						<tr>
							<td class="Item" style="WIDTH: 99px; HEIGHT: 31px" vAlign="top" align="right">研究領域及專長：</td>
							<td>
								<asp:Label id="LabelKMcat" runat="server" Font-Size="10">Label</asp:Label></td>
						</tr>
				</form>
				</TBODY></table>
		</form>
	</body>
</HTML>
