<%@ Page Language="C#" AutoEventWireup="true" MaintainScrollPositionOnPostback="true" CodeFile="Step3_EditUnit.aspx.cs" Inherits="GIP_web_Step3_EditUnit" %>
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
		<table cellspacing="0" class="setting" border="1">
			<caption>
				�i�ظm�y<asp:Label ID="TopicWebNameLabel" runat="server"></asp:Label>�z�����ؿ��j</caption>
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
							�ؿ���� �n
						</h5>
						<p>
							�ؿ��W�١G<asp:TextBox ID="NodeNameTextBox" runat="server" CssClass="box" Width="144px"></asp:TextBox>
							<asp:RequiredFieldValidator ID="NodeNameRequiredFieldValidator" runat="server" ControlToValidate="NodeNameTextBox" Display="Dynamic" ErrorMessage="�п�J�ؿ��W��"></asp:RequiredFieldValidator>
						</p>
						<p>
						    �ؿ������G<asp:TextBox ID="NodeNameMemoTextBox" runat="server" CssClass="box" Width="320px"></asp:TextBox>(�����|�X�{�b"��������"��)
						</p>
						<p>
							�O�_�}��G
							<asp:RadioButtonList ID="IsNodeOpenRadioButtonList" runat="server" RepeatDirection="Horizontal"
								RepeatLayout="Flow">
								<asp:ListItem Text="�O" Value="Y" Selected="True"></asp:ListItem>
								<asp:ListItem Text="�_" Value="N"></asp:ListItem>
							</asp:RadioButtonList>
						</p>
						<p>
							�ĤG�h�ؿ���ơG
							<asp:DropDownList ID="UnitTypeDropDownList" runat="server" CssClass="box" AutoPostBack="True" OnSelectedIndexChanged="UnitTypeDropDownList_SelectedIndexChanged">
								<asp:ListItem Text="�u�|���@�����(�u�ݭn���e��)" Value="CP"></asp:ListItem>
								<asp:ListItem Text="�|���ܦh�����(�ݭn�C���Τ��e��)" Value="LP"></asp:ListItem>
							</asp:DropDownList>
							<asp:RequiredFieldValidator ID="UnitTypeRequiredFieldValidator" runat="server" ControlToValidate="UnitTypeDropDownList" Display="Dynamic" ErrorMessage="�п�ܥؿ���ƫ��A"></asp:RequiredFieldValidator>
							<asp:Panel ID="ListTypePanel" runat="server" Width="100%" Visible="false">
								<fieldset>
									<legend>< �C���e�{�覡 > </legend>
									<p>
										��ܸ�����G
										<asp:CheckBoxList ID="ListFieldsCheckBoxList" runat="server" RepeatDirection="Horizontal"
											RepeatLayout="Flow" DataValueField="Key" DataTextField="Value">
										</asp:CheckBoxList>
									</p>
									<p>
										<div class="center">
											<asp:RadioButton ID="ListPageLayoutRadioButton1" runat="server" GroupName="ListPageLayoutRadioButton" /><br />
											<asp:Image ID="ListPageLayoutImage1" runat="server" ImageUrl="../images/lp01.gif"
												AlternateText="image" Width="120" />
										</div>
										<div class="center">
											<asp:RadioButton ID="ListPageLayoutRadioButton2" runat="server" GroupName="ListPageLayoutRadioButton" /><br />
											<asp:Image ID="ListPageLayoutImage2" runat="server" ImageUrl="../images/lp03.gif"
												AlternateText="image" Width="120" />
										</div>
										
									</p>
									<asp:Panel ID="BasicConfigPanel" runat="server" Width="100%" Visible="true">
										<asp:Button ID="AdvanceConfigButton" runat="server" CssClass="button rightbutton"
											Text="�i���]�w" OnClick="AdvanceConfigButton_Click" />
									</asp:Panel>
									<asp:Panel ID="AdvanceConfigPanel" runat="server" Width="100%" Visible="false">
										<hr />
										<p>
											��Ƥ����G
										</p>
										<p>
											<asp:DropDownList ID="DataCategoryTypeDropDownList" runat="server" CssClass="box" AppendDataBoundItems="true" DataTextField="Value" DataValueField="Key" AutoPostBack="True" OnSelectedIndexChanged="DataCategoryTypeDropDownList_SelectedIndexChanged">
												<asp:ListItem Value="" Text="���ݨϥΤ���"></asp:ListItem>
											</asp:DropDownList>
										</p>
										<asp:Panel ID="CategoryMaintainPanel" runat="server">
											<p>
												<asp:TextBox ID="CategoryNameTextBox" runat="server" CssClass="box" Width="110px"></asp:TextBox>
												<asp:Button ID="AddCategoryButton" runat="server" CssClass="button" Text="�s�W" ValidationGroup="DataCategory" OnClick="AddCategoryButton_Click" />
												<asp:Button ID="RemoveCategoryButton" runat="server" CssClass="button" Text="�R��" ValidationGroup="DataCategory" OnClick="RemoveCategoryButton_Click" />
												<asp:Button ID="CategoryMoveUpButton" runat="server" CssClass="button" Text="��" ValidationGroup="DataCategory" CausesValidation="false" OnClick="CategoryMoveUpButton_Click" />
												<asp:Button ID="CategoryMoveDownButton" runat="server" CssClass="button" Text="��" ValidationGroup="DataCategory" CausesValidation="false" OnClick="CategoryMoveDownButton_Click" />
											</p>
										</asp:Panel>
										<p>
											<asp:ListBox ID="DataCategoriesListBox" runat="server" CssClass="box" SelectionMode="Single" Width="232px" DataValueField="Key" DataTextField="Value"></asp:ListBox>
										</p>
										<p>
											<em>�`�N�I�t�δ��Ѫ��w�]�����A�����ѧR���C�Y�����U�w���D�D�峹�A�礣���ѧR���C</em></p>
										<asp:Button ID="BasicConfigButton" runat="server" CssClass="button rightbutton"
											Text="�򥻳]�w" OnClick="BasicConfigButton_Click" />
									</asp:Panel>
								</fieldset>
							</asp:Panel>
							<asp:Panel ID="ContentTypePanel" runat="server" Width="100%" Visible="false">
								<fieldset>
									<legend>&lt; ���e���e�{�覡 &gt; </legend>
									<p>
										��ܸ�����G
										<asp:CheckBoxList ID="ContentFieldsCheckBoxList" runat="server" RepeatDirection="Horizontal"
											RepeatLayout="Flow" DataValueField="Key" DataTextField="Value">
										</asp:CheckBoxList>
									</p>
									<p>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton1" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage1" runat="server" ImageUrl="../images/cp01.jpg"
												AlternateText="image" />
										</div>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton2" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage2" runat="server" ImageUrl="../images/cp02.jpg"
												AlternateText="image" />
										</div>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton3" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage3" runat="server" ImageUrl="../images/cp03.jpg"
												AlternateText="image" />
										</div>
									</p>
									<p>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton4" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage4" runat="server" ImageUrl="../images/cp04.jpg"
												AlternateText="image" />
										</div>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton5" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage5" runat="server" ImageUrl="../images/cp05.jpg"
												AlternateText="image" />
										</div>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton6" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage6" runat="server" ImageUrl="../images/cp06.jpg"
												AlternateText="image" />
										</div>
									</p>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton7" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage7" runat="server" ImageUrl="../images/cp07.jpg"
												AlternateText="image" />											
										</div>

								</fieldset>
							</asp:Panel>
                                    <p>
                                        <div class="center" id="NewsDiv" runat="server">
                                        ��L�]�w
                                        <asp:DropDownList ID="NewsDDL" runat="server" CssClass="box" AutoPostBack="True" OnSelectedIndexChanged="NewsDDL_SelectedIndexChanged">
                                            <asp:ListItem Text="�L�s�D�^���I�]�w" Value="N" Selected></asp:ListItem>
                                            <asp:ListItem Text="�]�w���s�D�^���I�]�w" Value="Y"></asp:ListItem>
                                        </asp:DropDownList>
                                        </div>
                                    </p>							
							<asp:Panel ID="NewsPanel" runat="server" Width="100%" Visible="false">
							<fieldset>
							<legend>&lt; �s�D�^���I�]�w &gt; </legend>
                                <table>
                                    <tr>
                                        <td>
                                        <fieldset>
                                        <legend> �ŦX����r  </legend>
                                            <asp:TextBox ID="addIntext" runat="server" Width="70px" ></asp:TextBox>
                                            <asp:Button ID="addInbutton" runat="server" Text="�[�J" OnClick="AddInButton_Click" Width="40px" />
                                            <asp:Button ID="delinbutton" runat="server" Text="����" OnClick="DelInButton_Click" Width="40px"/>
                                            <br />
                                            <asp:ListBox ID="NewsaddList" runat="server" Width="170px" SelectionMode="Single" DataValueField="Key" DataTextField="Value"></asp:ListBox>
                                        </fieldset></td>
                                        <td>
                                        <fieldset>
                                        <legend> �ư�����r  </legend>
                                            <asp:TextBox ID="addOuttext" runat="server" Width="70px"></asp:TextBox>
                                            <asp:Button ID="addoutbutton" runat="server" Text="�[�J" OnClick="AddOutButton_Click" Width="40px"/>
                                            <asp:Button ID="deloutbutton" runat="server" Text="����" OnClick="DelOutButton_Click" Width="40px" />
                                            <br />
                                            <asp:ListBox ID="NewsOrList" runat="server" Width="170px" SelectionMode="Single" DataValueField="Key" DataTextField="Value"></asp:ListBox>
                                        </fieldset></td>
                                    </tr>
                                </table>																		
                                </fieldset>
							</asp:Panel>							
							<div class="settingbutton">
								<asp:Button ID="UpdateButton" runat="server" Text="�s�צs��" OnClick="UpdateButton_Click" />
								<asp:Button ID="DeleteButton" runat="server" Text="�R��" OnClientClick="return confirm('�T�w�n�R���H');" OnClick="DeleteButton_Click" />
								<input type="reset" name="ResetButton" value="����" />
								<asp:Button ID="cancel" runat="server" Text="����" OnClick="cancel_Click" />
							</div>
						</p>
					</asp:Panel>
				</td>
			</tr>
		</table>
	</form>
</body>
</html>