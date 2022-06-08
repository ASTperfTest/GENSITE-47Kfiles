<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Step3_AddFolder.aspx.cs" Inherits="GIP_web_Step3_AddFolder" %>
<%@ Register TagPrefix="gip" TagName="CatelogTree" Src="~/GIP/web/CatalogTreeUserControl.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<link href="../css/list.css" rel="stylesheet" type="text/css">
	<link href="../css/layout.css" rel="stylesheet" type="text/css">
	<link href="../css/theme.css" rel="stylesheet" type="text/css">
	<title>主題館管理</title>
</head>
<body>
	<form id="form1" runat="server">
		<div id="FuncName">
			<h1>
				新增主題館</h1>
			<div id="ClearFloat">
			</div>
		</div>
		<div class="step">
			步驟： 1. 填寫基本資料 > 2. 設定版面風格 > <span>3. 導覽架構設定</span> > 4. 首頁配置</div>
		<table cellspacing="0" class="setting">
			<caption>
				【建置『<asp:Label ID="TopicWebNameLabel" runat="server"></asp:Label>』導覽目錄】</caption>
			<tr>
				<th width="30%">
				</th>
				<th>
					<div class="right">
					</div>
				</th>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF">
					<gip:CatelogTree ID="CatelogTree1" runat="server" />
				</td>
				<td>
					<asp:Panel ID="MainPanel" runat="server">
						<h5>
							目錄資料 》</h5>
						<p>
							目錄名稱：<asp:TextBox ID="FolderNameTextBox" runat="server" CssClass="box" Width="144px"></asp:TextBox>
						</p>
						<p>
						    目錄說明：<asp:TextBox ID="NodeNameMemoTextBox" runat="server" CssClass="box" Width="144px" ></asp:TextBox>
							&nbsp;&nbsp;(說明會出現在"網站導覽"中)
						</p>
						<p>
							是否開放：
							<asp:RadioButtonList ID="IsFolderOpenRadioButtonList" runat="server" RepeatDirection="Horizontal"
								RepeatLayout="Flow">
								<asp:ListItem Text="是" Value="Y" Selected="True"></asp:ListItem>
								<asp:ListItem Text="否" Value="N"></asp:ListItem>
							</asp:RadioButtonList>
						</p>
						<p>
							是否有下一層：
							<asp:RadioButtonList ID="HasChildRadioButtonList" runat="server" RepeatDirection="Horizontal"
								RepeatLayout="Flow">
								<asp:ListItem Text="是" Value="Y" Selected="True"></asp:ListItem>
								<asp:ListItem Text="否" Value="N"></asp:ListItem>
							</asp:RadioButtonList>
						</p>
						<div class="settingbutton">
							<asp:Button ID="InsertButton" runat="server" Text="確定新增" OnClick="InsertButton_Click" />
							<asp:Button ID="ResetButton" runat="server" Text="重填" OnClientClick="form1.reset();" UseSubmitBehavior="false" />
							<asp:Button ID="cancelButton" runat="server" Text="取消" OnClick="cancelButton_Click" />
						</div>
					</asp:Panel>
				</td>
			</tr>
		</table>
	</form>
</body>
</html>