<%@ Page Language="C#" AutoEventWireup="true" CodeFile="class_node_edit.aspx.cs" Inherits="Edit_class_node_edit" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>未命名頁面</title>
 <link href="../css/list.css" rel="stylesheet" type="text/css" />
    <link href="../css/layout.css" rel="stylesheet" type="text/css"/>
    <link href="../css/theme.css" rel="stylesheet" type="text/css"/>
     <script language="javascript" type="text/javascript">
     function check_name()
{
      
     var Title = document.getElementById("Txt_Class");
     var sort = document.getElementById("Txt_sortvalue");
     var type = document.getElementById("ddl_class1");
    
    if(type.value == "--請選擇--")
	{
	    alert("請選擇分類")
	    event.returnValue = false;
	} 
	else if(Title.value == "")
	{	
		alert("請輸入類別名稱");
		event.returnValue = false;
		
	}
	else if(sort.value == "")
	{
	    alert("請輸入排序值(1~99)")
	    event.returnValue = false;
	}
	else if(isNaN(sort.value))
	{
	    alert("排序值請輸入數字!(1~99)")
	    event.returnValue = false;
	}

	
	
}
</script>
</head>
<body style="text-align: center">
  <div id="FuncName">
	    <h1>分類管理</h1>
	    <div id="ClearFloat"></div>
    </div>
    <form id="form1" runat="server">
    <div style="text-align: center">
        <asp:Label ID="Label2" runat="server" Text="父類別名稱" Visible="False"></asp:Label>
        <asp:DropDownList ID="ddl_class1" runat="server" Visible="False">
        </asp:DropDownList>
        &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        <br />
        <asp:Label ID="Label1" runat="server" Text="類別名稱" Visible="False"></asp:Label>
        <asp:TextBox ID="Txt_Class" runat="server" Visible="False"></asp:TextBox><br />
        
        <asp:Label ID="Sort" runat="server" Text="顯示順序"></asp:Label>
        <asp:TextBox ID="Txt_sortvalue" runat="server" Width="30px" MaxLength="2"></asp:TextBox>(1~99) &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        
    </div>
   
    <br />
    <asp:Button ID="btn_Add" runat="server" Text="新增" OnClientClick="check_name()" OnClick="btn_Add_Click" Visible="False" />
    <asp:Button ID="btn_Edit" runat="server" Text="確定維護" OnClientClick="check_name()" OnClick="btn_Edit_Click" Visible="False" />
    <asp:Button ID="btn_Cancel" runat="server" Text="取消" OnClick="btn_Cancel_Click" Visible="False" />
    <asp:Button ID="btn_Delete" runat="server" Text="刪除" OnClientClick="return confirm('確定刪除？')" OnClick="btn_Delete_Click" Visible="False" />&nbsp;<br />
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    &nbsp; &nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
        &nbsp; &nbsp; &nbsp;<br />
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
    &nbsp; &nbsp; &nbsp; &nbsp;&nbsp;&nbsp;
        <asp:Label ID="lbl_worng" runat="server" ForeColor="Red" Text="該分類尚有子項資料，無法刪除！" Font-Size="X-Small" Visible="False"></asp:Label>
        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<br />
    </form>
</body>
</html>
