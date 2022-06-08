<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Step3.aspx.cs" Inherits="GIP_web_Step3" %>
<%@ Register TagPrefix="gip" TagName="CatelogTree" Src="~/GIP/web/CatalogTreeUserControl.ascx" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
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
						<asp:ImageButton ID="AddCatelogFolderImageButton" runat="server" ImageUrl="../images/bt_addmenu1.gif" OnClick="AddCatelogFolderImageButton_Click" />
					</div>
				</th>
			</tr>
			<tr>
				<td bgcolor="#FFFFFF">
					<gip:CatelogTree ID="CatelogTree1" runat="server" />
				</td>
				<td>
					<p>�i�I��s�W�Ĥ@�h�ؿ��ؿ��W�s����r�A�s�W�D�D�]�ؿ��[�c�A �άO�I�索��ؿ��[�c�W�١A�s�שҿ�����ؿ���ơC</p>
					<p>�s�W���ؿ��N�|�m��ؿ��[�c���U��C</p>
					<p>�ؿ������|�X�{�b�D�D�]�� ������������ ���C</p>
				</td>
			</tr>
		</table>
		<div class="settingbutton">
			<asp:Button ID="PrevStepButton" runat="server" Text="�W�@�B" OnClientClick="history.back();" OnClick="PrevStepButton_Click" />
			<asp:Button ID="NextStepButton" runat="server" Text="�U�@�B" OnClick="NextStepButton_Click" />
			<asp:Button ID="SaveButton" runat="server" Text="�x�s�]�w" Visible="false" />
			<asp:Button ID="CancelButton" runat="server" Text="���� (�^����)" OnClick="CancelButton_Click" />
		</div>
	</form>
</body>
</html>
