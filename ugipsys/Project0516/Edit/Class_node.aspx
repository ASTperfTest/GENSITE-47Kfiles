<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Class_node.aspx.cs" Inherits="Edit_Class_node" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>未命名頁面</title>
    <link href="../css/list.css" rel="stylesheet" type="text/css" />
    <link href="../css/layout.css" rel="stylesheet" type="text/css"/>
    <link href="../css/theme.css" rel="stylesheet" type="text/css"/>

</head>
<body style="text-align: center">
  <div id="FuncName">
	    <h1>分類管理</h1>
	    <div id="ClearFloat"></div>
    </div>
    <form id="form1" runat="server">

    
    <div style="text-align: center">
       
        <asp:HyperLink ID ="goback" runat="server" NavigateUrl="~/edit/class.aspx" text="回上頁" CssClass="newsite"></asp:HyperLink>

        <asp:HyperLink ID ="link_add" runat="server" NavigateUrl="~/edit/class_node_edit.aspx?id=0" text="新增子分類" CssClass="newsite"></asp:HyperLink>
        
        <br />
        <br />
         <asp:Label ID="Label_Class" runat="server" Text=""></asp:Label>
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            CssClass="ListTable" DataSourceID="SqlDataSource1" HorizontalAlign="Center" PageSize="20" OnDataBound="GridView1_DataBound">
            <PagerSettings Visible="False" />
            <Columns>
                <asp:BoundField DataField="classid" HeaderText="ID" />
                <asp:BoundField DataField="classname" HeaderText="子分類名稱" />
                <asp:BoundField HeaderText="分類管理" />
            </Columns>
        </asp:GridView>
        &nbsp;
        
    <br />
    <br />
    <br />
    <br />
    <asp:SqlDataSource ID="SqlDataSource1" runat="server" OnSelecting="SqlDataSource1_Selecting"></asp:SqlDataSource>
     </form>
</body>
</html>
