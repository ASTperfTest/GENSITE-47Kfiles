<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Finish.aspx.cs" Inherits="Finish" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
<link href="./css/list.css" rel="stylesheet" type="text/css"/>
<link href="./css/layout.css" rel="stylesheet" type="text/css"/>
<link href="./css/theme.css" rel="stylesheet" type="text/css"/>
    <title>未命名頁面</title>
</head>
<body>
    <form id="form1" runat="server">
    <div id="FuncName">
	<h1>新增主題館</h1>
	<div id="ClearFloat"></div>
</div>

<div class="setting" align="center">
<h5>您已成功新增主題館 ! </h5>
<p>請按『完成』按鈕，來關閉這個精靈，或點選『維護主題文章』按鈕，接續進行主題文章維護</p>
</div>

<div class="settingbutton">
<asp:Button ID="browser" runat="server" Text="主題館預覽" />
<%--<input name="Submit" type="submit" onClick="MM_openBrWindow('../home_style/style03.htm','','toolbar=yes,location=yes,status=yes,menubar=yes,scrollbars=yes,width=980,height=600')" value="主題館預覽" id="Submit1">
--%><%--<input name="button" type="button" onClick="location.href='../new_article/new_article_new.htm'" value="維護主題文章">
<input name="button" type="button" onClick="location.href='new_web_list.htm'" value="完成(回首頁)">
--%>
<asp:Button ID="modify" runat="server" Text="維護主題文章" OnClick="modify_Click" />
<asp:Button ID="gototop" runat="server" Text="完成(回首頁)" OnClick="gototop_Click"/>
</div>
    </form>
</body>
</html>
