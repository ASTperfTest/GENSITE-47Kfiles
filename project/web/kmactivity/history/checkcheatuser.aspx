<%@ Page Language="C#" AutoEventWireup="true" CodeFile="checkcheatuser.aspx.cs" Inherits="kmactivity_history_checkcheatuser" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../js/jquery-1.3.2.min.js"></script>

    <script type="text/javascript" src="../js/datepicker.js"></script>

    <link rel="Stylesheet" href="../css/datepicker.css" />
    <link type="text/css" href="../css/jquery.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
    <div class="content" style="padding-left:20px; padding-top:20px;">
      <input type="button" onclick="javascript:history.go(-1);" value="回上一頁" /><br />
      每日答題超過30題清單&nbsp;&nbsp;排序依<asp:DropDownList Font-Size="Smaller" ID="DropDownList1" runat="server" AutoPostBack="true">
        <asp:ListItem Value="0">日期</asp:ListItem>  
        <asp:ListItem Value="1">帳號</asp:ListItem> 
        </asp:DropDownList>
        <br />
      <asp:Label ID="TableText" runat="server" Text="" />  <br />
      每日得分超過180分的使用者&nbsp;&nbsp;排序依<asp:DropDownList Font-Size="Smaller" ID="DropDownList2" runat="server" AutoPostBack="true">
        <asp:ListItem Value="0">日期</asp:ListItem>  
        <asp:ListItem Value="1">帳號</asp:ListItem> 
        </asp:DropDownList><br />
      <asp:Label ID="Label1" runat="server" Text="" />  <br />
      單次得分超過6分的使用者&nbsp;&nbsp;排序依<asp:DropDownList Font-Size="Smaller" ID="DropDownList3" runat="server" AutoPostBack="true">
        <asp:ListItem Value="0">日期</asp:ListItem>  
        <asp:ListItem Value="1">帳號</asp:ListItem> 
        </asp:DropDownList><br />
      <asp:Label ID="Label2" runat="server" Text="" />  
    </div>
    </form>
</body>
</html>