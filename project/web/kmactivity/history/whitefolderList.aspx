<%@ Page Language="C#" AutoEventWireup="true" CodeFile="whitefolderList.aspx.cs" Inherits="kmactivity_history_whitefolderList" %>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="css/style.css" rel="stylesheet" type="text/css" />
    <link rel="Stylesheet" href="/css/datepicker.css" />
    <link type="text/css" href="/css/jquery.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        加入節點：<asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="確定" />
        <hr size="1">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" OnRowDeleting="GridView1_RowDeleting" DataKeyNames="NodeId">
            <Columns>
                <asp:BoundField DataField="NodeId" HeaderText="節點代碼" />
                <asp:BoundField DataField="FullPath" HeaderText="完整路徑" />
                <asp:CommandField ShowDeleteButton="True" />
            </Columns>
        </asp:GridView>
    </div>
    </form>
    
    <script type="text/javascript">
        var msg = '<%=errmsg %>'
        if (msg != '') alert(msg);
    </script>
</body>
</html>
