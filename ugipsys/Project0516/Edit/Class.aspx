<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Class.aspx.cs" Inherits="Edit_Class" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
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

         <asp:HyperLink ID ="link_add" runat="server" NavigateUrl="~/edit/class_edit.aspx?id=0" text="新增分類" CssClass="newsite"></asp:HyperLink>

    <div style="text-align: center">
        <br />
        <br />
        <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
            CssClass="ListTable" DataSourceID="SqlDataSource1" HorizontalAlign="Center" PageSize="20">
            <PagerSettings Visible="False" />
            <Columns>
                <asp:BoundField DataField="classid" HeaderText="ID" Visible="False" />
                <asp:HyperLinkField DataNavigateUrlFields="classid,classname" DataNavigateUrlFormatString="class_edit.aspx?id={0}"
                    DataTextField="classname" DataTextFormatString="{0}" HeaderText="分類管理" NavigateUrl="class_edit.aspx?id={0}" />
                <asp:HyperLinkField DataNavigateUrlFields="classid,classname" DataNavigateUrlFormatString="class_node.aspx?id={0}"
                    DataTextFormatString="{0}" HeaderText="子分類管理" NavigateUrl="class_node.aspx?id={0}"
                    Text="管理" />
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
