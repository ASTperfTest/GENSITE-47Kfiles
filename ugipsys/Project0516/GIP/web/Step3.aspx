<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Step3.aspx.cs" Inherits="GIP_web_Step3" %>
<%@ Register TagPrefix="gip" TagName="CatelogTree" Src="~/GIP/web/CatalogTreeUserControl.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
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
			步驟： 1. <a href="../../edit/step1_edit.aspx">填寫基本資料</a> > 2. <a href="../../edit/step2_edit.aspx">設定版面風格</a> > <span>3. 導覽架構設定</span> > 4. <a href="../../HomePageEdit.aspx">首頁配置</a></div>
		<table cellspacing="0" class="setting">
			<caption>
				【建置『<asp:Label ID="TopicWebNameLabel" runat="server"></asp:Label>』導覽目錄】</caption>
			<tr>
				<th width="30%">
				</th>
				<th>
					<div class="right">
						<asp:ImageButton ID="AddCatelogFolderImageButton" runat="server" ImageUrl="../images/bt_addmenu1.gif" OnClick="AddCatelogFolderImageButton_Click" />
					</div>
				</th>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF">
					<gip:CatelogTree ID="CatelogTree1" runat="server" />
				</td>
				<td>
					<p>可點選新增第一層目錄目錄超連結文字，新增主題館目錄架構， 或是點選左方目錄架構名稱，編修所選取的目錄資料。</p>
					<p>新增的目錄將會置於目錄架構的下方。</p>
					<p>目錄說明會出現在主題館的 “網站導覽” 中。</p>
				</td>
			</tr>
		</table>
		<div class="settingbutton">
			<asp:Button ID="PrevStepButton" runat="server" Text="上一步" OnClientClick="history.back();" OnClick="PrevStepButton_Click" />
			<asp:Button ID="NextStepButton" runat="server" Text="下一步" OnClick="NextStepButton_Click" />
			<asp:Button ID="SaveButton" runat="server" Text="儲存設定" Visible="false" />
			<asp:Button ID="CancelButton" runat="server" Text="取消 (回首頁)" OnClick="CancelButton_Click" />
		</div>
	</form>
</body>
</html>
