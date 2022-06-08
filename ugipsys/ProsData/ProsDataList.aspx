<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ProsDataList.aspx.vb" Inherits="ProsData_ProsDataList" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <link href="/css/list.css" rel="stylesheet" type="text/css">
    <link href="/css/layout.css" rel="stylesheet" type="text/css">
</head>
<body>
    <div id="FuncName">
	    <h1>專家資料／專家列表</h1><font size="2">【目錄樹節點: 專家列表】</font>
	    <div id="Nav"></div>
	    <div id="ClearFloat"></div>
    </div>
    <div id="FormName">
    	單元資料維護&nbsp;<font size="2">【主題單元:專家列表 / 單元資料:純網頁】</font>
    </div>
    <form id="Form2" name="reg" runat="server">    
        <!-- 分頁 -->
        <div id="Page">
             
		    共 <font color="red"><asp:Label ID="TotalCountTxt" runat="server" Text="0"></asp:Label></font> 筆資料，每頁顯示
		    
            <asp:DropDownList ID="PageSizeDDL" runat="server" AutoPostBack="true">
                <asp:ListItem>15</asp:ListItem>
		        <asp:ListItem>30</asp:ListItem>
		        <asp:ListItem>50</asp:ListItem>
		        <asp:ListItem>300</asp:ListItem>		    
            </asp:DropDownList>		    		    
       		筆，目前在第
            <asp:DropDownList ID="PageNumberDDL" runat="server" AutoPostBack="true"></asp:DropDownList>            
		    頁，		    
            <img src="/images/arrow_previous.gif" >
            <asp:LinkButton ID="BackLBtn" runat="server" Text="上一頁" CommandArgument="Prev" ></asp:LinkButton>
            <asp:LinkButton ID="NextLBtn" runat="server" Text="下一頁" CommandArgument="Next" ></asp:LinkButton>
            <img src="/images/arrow_next.gif" alt="下一頁">
        </div>

        <asp:GridView ID="ListTable" runat="server" DataSourceID="SqlDataSource1" AutoGenerateColumns="False" AllowPaging="True" PageSize="15" 
             CssClass="ListTable" OnPageIndexChanging="ListTable_PageIndexChanging" >
            <Columns>
            
                <asp:BoundField DataField="account" HeaderText="專家帳號" Visible="False"/>                             
                <asp:HyperLinkField DataTextField="realname" HeaderText="專家姓名"  NavigateUrl="ProsDataDetail.aspx?account={0}" DataNavigateUrlFields="account" 
                DataNavigateUrlFormatString="ProsDataDetail.aspx?account={0}" HeaderStyle-CssClass="First" ItemStyle-CssClass="eTableContent" />                               
                <asp:BoundField DataField="nickname" HeaderText="專家暱稱" />
            </Columns>
            <PagerSettings Visible="False" />
        </asp:GridView>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server"></asp:SqlDataSource>
    </form>
</body>
</html>
