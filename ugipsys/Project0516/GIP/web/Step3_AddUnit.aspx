<%@ Page Language="C#" AutoEventWireup="true" MaintainScrollPositionOnPostback="true" CodeFile="Step3_AddUnit.aspx.cs" Inherits="GIP_web_Step3_AddUnit" %>
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
						<asp:ImageButton ID="AddCatelogFolderImageButton" runat="server" ImageUrl="../images/bt_addmenu1.gif" OnClick="AddCatelogFolderImageButton_Click" />
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
							目錄資料 》
						</h5>
						<p>
							目錄名稱：<asp:TextBox ID="NodeNameTextBox" runat="server" CssClass="box" Width="144px"></asp:TextBox>
							<asp:RequiredFieldValidator ID="NodeNameRequiredFieldValidator" runat="server" ControlToValidate="NodeNameTextBox" Display="Dynamic" ErrorMessage="請輸入目錄名稱"></asp:RequiredFieldValidator>
						</p>
						<p>
						    目錄說明：<asp:TextBox ID="NodeNameMemoTextBox" runat="server" CssClass="box" Width="144px"></asp:TextBox>
							&nbsp;&nbsp;(說明會出現在"網站導覽"中)
						</p>
						<p>
							是否開放：
							<asp:RadioButtonList ID="IsNodeOpenRadioButtonList" runat="server" RepeatDirection="Horizontal"
								RepeatLayout="Flow">
								<asp:ListItem Text="是" Value="Y" Selected="True"></asp:ListItem>
								<asp:ListItem Text="否" Value="N"></asp:ListItem>
							</asp:RadioButtonList>
						</p>
						<p>
							第二層目錄資料：
							<asp:DropDownList ID="UnitTypeDropDownList" runat="server" CssClass="box" AutoPostBack="True" OnSelectedIndexChanged="UnitTypeDropDownList_SelectedIndexChanged">
								<asp:ListItem Text="請選擇" Value=""></asp:ListItem>
								<asp:ListItem Text="只會有一筆資料(只需要內容頁)" Value="CP"></asp:ListItem>
								<asp:ListItem Text="會有很多筆資料(需要列表頁及內容頁)" Value="LP"></asp:ListItem>
							</asp:DropDownList>
							<asp:RequiredFieldValidator ID="UnitTypeRequiredFieldValidator" runat="server" ControlToValidate="UnitTypeDropDownList" Display="Dynamic" ErrorMessage="請選擇目錄資料型態"></asp:RequiredFieldValidator>
							<asp:Panel ID="ListTypePanel" runat="server" Width="100%" Visible="false">
								<fieldset>
									<legend>< 列表頁呈現方式 > </legend>
									<p>
										顯示資料欄位：
										<asp:CheckBoxList ID="ListFieldsCheckBoxList" runat="server" RepeatDirection="Horizontal"
											RepeatLayout="Flow" DataValueField="Key" DataTextField="Value">
										</asp:CheckBoxList>
									</p>
									<p>
										<div class="center">
											<asp:RadioButton ID="ListPageLayoutRadioButton1" runat="server" GroupName="ListPageLayoutRadioButton" Checked="True" /><br />
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
										<asp:Button ID="AdvanceConfigButton" runat="server" CssClass="button rightbutton" CausesValidation="false"
											Text="進階設定" OnClick="AdvanceConfigButton_Click" />
									</asp:Panel>
									<asp:Panel ID="AdvanceConfigPanel" runat="server" Width="100%" Visible="false">
										<hr />
										<p>
											資料分類：
										</p>
										<p>
											<asp:DropDownList ID="DataCategoryTypeDropDownList" runat="server" CssClass="box" AppendDataBoundItems="true" DataTextField="Value" DataValueField="Key" AutoPostBack="True" OnSelectedIndexChanged="DataCategoryTypeDropDownList_SelectedIndexChanged">
												<asp:ListItem Value="" Text="不需使用分類"></asp:ListItem>
											</asp:DropDownList>
										</p>
										<asp:Panel ID="CategoryMaintainPanel" runat="server">
											<p>
												<asp:TextBox ID="CategoryNameTextBox" runat="server" CssClass="box" Width="110px"></asp:TextBox>
												<asp:Button ID="AddCategoryButton" runat="server" CssClass="button" Text="新增" ValidationGroup="DataCategory" OnClick="AddCategoryButton_Click" />
												<asp:Button ID="RemoveCategoryButton" runat="server" CssClass="button" Text="刪除" ValidationGroup="DataCategory" OnClick="RemoveCategoryButton_Click" />
												<asp:Button ID="CategoryMoveUpButton" runat="server" CssClass="button" Text="↑" ValidationGroup="DataCategory" CausesValidation="false" OnClick="CategoryMoveUpButton_Click" />
												<asp:Button ID="CategoryMoveDownButton" runat="server" CssClass="button" Text="↓" ValidationGroup="DataCategory" CausesValidation="false" OnClick="CategoryMoveDownButton_Click" />
											</p>
										</asp:Panel>
										<p>
											<asp:ListBox ID="DataCategoriesListBox" runat="server" CssClass="box" SelectionMode="Single" Width="232px" DataValueField="Key" DataTextField="Value"></asp:ListBox>
										</p>
										<p>
											<em>注意！系統提供的預設分類，不提供刪除。若分類下已有主題文章，亦不提供刪除。</em></p>
										<asp:Button ID="BasicConfigButton" runat="server" CssClass="button rightbutton" CausesValidation="false"
											Text="基本設定" OnClick="BasicConfigButton_Click" />
									</asp:Panel>
								</fieldset>
							</asp:Panel>
							<asp:Panel ID="ContentTypePanel" runat="server" Width="100%" Visible="false">
								<fieldset>
									<legend>&lt; 內容頁呈現方式 &gt; </legend>
									<p>
										顯示資料欄位：
										<asp:CheckBoxList ID="ContentFieldsCheckBoxList" runat="server" RepeatDirection="Horizontal"
											RepeatLayout="Flow" DataValueField="Key" DataTextField="Value">
										</asp:CheckBoxList>
									</p>
									<p>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton1" runat="server" GroupName="ContentPageLayoutRadioButton" Checked="True" /><br />
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
									<p>
										<div class="center">
											<asp:RadioButton ID="ContentPageLayoutRadioButton7" runat="server" GroupName="ContentPageLayoutRadioButton" /><br />
											<asp:Image ID="ContentPageLayoutImage7" runat="server" ImageUrl="../images/cp07.jpg"
												AlternateText="image" />											
										</div>									
									</p>
								</fieldset>
							</asp:Panel>
							<div class="settingbutton">
								<asp:Button ID="InsertButton" runat="server" Text="確定新增" OnClick="InsertButton_Click" />
								<input type="reset" name="ResetButton" value="重填" />
								<asp:Button ID="cancel" runat="server" Text="取消" OnClick="cancel_Click" />
							</div>
						</p>
					</asp:Panel>
				</td>
			</tr>
		</table>
	</form>
</body>
</html>
