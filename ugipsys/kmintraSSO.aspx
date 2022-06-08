<%@ Page Language="C#" AutoEventWireup="true" CodeFile="kmintraSSO.aspx.cs" Inherits="kmintraSSO" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>農業知識管理網絡系統</title>
</head>
<body style="background-color: #ebf5e2; font-size: 13px">
    <form id="form1" runat="server" onsubmit="window.close()">
    <div style="padding:20px 20px 20px 20px; margin-left:auto; margin-top:auto;">
        請輸入查詢關鍵字：
        <input type="text" id="searchword" runat="server" value="" name="intraKW" onkeypress="if(window.event.keyCode==13) return false;" />
        <asp:Button ID="btnSearch" runat="server" Text="查詢" onclick="btnSearch_Click" />
        <input type="button" onclick="window.close()" value="離開" />
    </div>
    </form>
</body>
</html>
