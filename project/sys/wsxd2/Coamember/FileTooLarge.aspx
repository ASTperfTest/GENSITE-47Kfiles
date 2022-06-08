<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FileTooLarge.aspx.cs" Inherits="FileTooLarge" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <script type="text/javascript">
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
		<br />
		<br />
		
        <asp:Label ID="LabelErrorMsg" runat="server" ForeColor="red" Font-Size="Small"></asp:Label>
       
        &nbsp;<asp:Button ID="BackButton" runat="server" Text="回上傳畫面" OnClick="BackButton_Click"  />
    
    </div>
    </form>
</body>
</html>
