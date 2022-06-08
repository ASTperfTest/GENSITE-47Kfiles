<%@ Page Language="C#" AutoEventWireup="true" CodeFile="HomePageEdit.aspx.cs" Inherits="HomePageEdit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <link href="css/list.css" rel="stylesheet" type="text/css" />
    <link href="css/layout.css" rel="stylesheet" type="text/css" />
    <link href="css/theme.css" rel="stylesheet" type="text/css" />
    <link href="css/Display.css" rel="stylesheet" type="text/css" />
    <title>主題館管理</title>
</head>
<body>

<div id="FuncName">
	<h1>新增主題館</h1>
	<div id="ClearFloat"></div>
</div>

<div class="step">步驟： 1. <a href="edit/step1_edit.aspx">填寫基本資料</a> > 2. <a href="edit/step2_edit.aspx">設定版面風格</a> > 3. <a href="GIP/web/step3.aspx">導覽架構設定</a> > <span>4. 首頁配置</span></div>

<form id="form1" runat="server">
<!--列表-->
    <table cellspacing="0" class="setting">
    <caption>【設定主題館首頁配置】(請依據版面配置方式，選取各編號區塊的顯示方式)</caption>
    
    <tr>
        <td>
            <h5>您選取的版面配置 》</h5>
            <asp:ImageButton ID="LayoutPic" runat="server" />    
        </td>
                
        <td>
            <h5>區塊顯示資料＆顯示方式設定 》</h5>
            <asp:Panel ID="Panel1" runat="server">
            </asp:Panel>
                  
        </td>
    </tr>
    </table>
    
    <div class="settingbutton">
        <asp:Button ID="Back" runat="server" Text="上一步" OnClick="Back_Click" />
        <asp:Button ID="Next" runat="server" Text="下一步" OnClick="Next_Click" />
        <asp:Button ID="Save" runat="server" Text="儲存設定" OnClick="Save_Click" />
        <asp:Button ID="Home" runat="server" Text="取消(回首頁)" OnClick="Home_Click" />
</div>
   
</form>
                            
</body>
</html>