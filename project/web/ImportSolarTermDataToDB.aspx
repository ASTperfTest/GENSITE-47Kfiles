<%@ Page Language="C#" AutoEventWireup="true" CodeFile="ImportSolarTermDataToDB.aspx.cs" Inherits="ImportSolarTermDataToDB" Debug="true"%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>新增節氣資料</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
	<asp:Label ID="Label1" runat="server"></asp:Label>
        <br />
        <hr />
        &nbsp;</div>自<asp:Label ID="Label2" runat="server"></asp:Label>起
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="塞20年節氣資料到DB" />
		
    </form>
</body>
</html>
