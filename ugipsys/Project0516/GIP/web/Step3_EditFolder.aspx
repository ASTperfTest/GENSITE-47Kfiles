<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Step3_EditFolder.aspx.cs" Inherits="GIP_web_Step3_EditFolder" %>
<%@ Register TagPrefix="gip" TagName="CatelogTree" Src="~/GIP/web/CatalogTreeUserControl.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
	<meta http-equiv="Content-Type" content="text/html; charset=big5">
	<link href="../css/list.css" rel="stylesheet" type="text/css">
	<link href="../css/layout.css" rel="stylesheet" type="text/css">
	<link href="../css/theme.css" rel="stylesheet" type="text/css">
	<title>�D�D�]�޲z</title>
</head>
<body>
	<form id="form1" runat="server">
		<div id="FuncName">
			<h1>
				�s�W�D�D�]</h1>
			<div id="ClearFloat">
			</div>
		</div>
		<div class="step">
			�B�J�G 1. <a href="../../edit/step1_edit.aspx">��g�򥻸��</a> > 2. <a href="../../edit/step2_edit.aspx">�]�w��������</a> > <span>3. �����[�c�]�w</span> > 4. <a href="../../HomePageEdit.aspx">�����t�m</a></div>
		<table cellspacing="0" class="setting">
			<caption>
				�i�ظm�y<asp:Label ID="TopicWebNameLabel" runat="server"></asp:Label>�z�����ؿ��j</caption>
			<tr>
				<th width="30%">
				</th>
				<th>
					<div class="right">
						<asp:ImageButton ID="AddCatelogNodeImageButton" runat="server" ImageUrl="../images/bt_addmenu2.gif" OnClick="AddCatelogNodeImageButton_Click" />
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
							�ؿ���� �n</h5>
						<p>
							�ؿ��W�١G<asp:TextBox ID="FolderNameTextBox" runat="server" CssClass="box" Width="144px"></asp:TextBox>
						</p>
						<p>
						    �ؿ������G<asp:TextBox ID="NodeNameMemoTextBox" runat="server" CssClass="box" Width="144px"></asp:TextBox>
							&nbsp;&nbsp;(�����|�X�{�b"��������"��)
						</p>
						<p>
							�O�_�}��G
							<asp:RadioButtonList ID="IsFolderOpenRadioButtonList" runat="server" RepeatDirection="Horizontal"
								RepeatLayout="Flow">
								<asp:ListItem Text="�O" Value="Y" Selected="True"></asp:ListItem>
								<asp:ListItem Text="�_" Value="N"></asp:ListItem>
							</asp:RadioButtonList>
						</p>
						<div class="settingbutton">
							<asp:Button ID="UpdateButton" runat="server" Text="�s�צs��" OnClick="UpdateButton_Click" />
							<asp:Button ID="DeleteButton" runat="server" Text="�R��" OnClientClick="return confirm('�T�w�n�R���H');" OnClick="DeleteButton_Click" />
							<asp:Button ID="ResetButton" runat="server" Text="����" OnClientClick="form1.reset();" UseSubmitBehavior="false" />
							<asp:Button ID="Cancel" runat="server" Text="����" OnClick="Cancel_Click" />
						</div>
					</asp:Panel>
				</td>
			</tr>
		</table>
	</form>
</body>
</html>