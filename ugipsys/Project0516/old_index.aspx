<%@ Page Language="C#" AutoEventWireup="true" CodeFile="old_index.aspx.cs" Inherits="old_index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
  <head runat="server">
    <link href="./css/list.css" rel="stylesheet" type="text/css"/>
    <link href="./css/layout.css" rel="stylesheet" type="text/css"/>
    <link href="./css/theme.css" rel="stylesheet" type="text/css"/>
    <title>主題館封存區</title>
  </head>
  <body>
    <div id="FuncName">
	    <h1>主題館封存區</h1>
	    <div id="ClearFloat"></div>
    </div>
    <form id="form1" runat="server">
            
      <div id="Page">             
		    共<asp:Label ID="datacount" runat="server" Text="0"></asp:Label>筆資料，每頁顯示
		    <asp:DropDownList ID="pagesize" runat="server" OnSelectedIndexChanged="pagesize_SelectedIndexChanged" AutoPostBack="True" >
		      <asp:ListItem>15</asp:ListItem>
		      <asp:ListItem>30</asp:ListItem>
		      <asp:ListItem>50</asp:ListItem>
		      <asp:ListItem>300</asp:ListItem>
		    </asp:DropDownList>
        筆，目前在第
        <asp:DropDownList ID="ddl_page" runat="server" OnSelectedIndexChanged="ddl_page_SelectedIndexChanged" AutoPostBack="True"></asp:DropDownList>
		    頁，
        <img src="./images/arrow_previous.gif" alt="" />
        <asp:LinkButton ID="back" runat="server" Text="上一頁" CommandArgument="Prev" OnClick="back_Click"></asp:LinkButton>
        <asp:LinkButton ID="next" runat="server" Text="下一頁" CommandArgument="Next" OnClick="next_Click"></asp:LinkButton>
        <img src="./images/arrow_next.gif" alt="下一頁" />
        
      </div>

      <asp:GridView ID="GridView1" runat="server" DataSourceID="SqlDataSource1" AutoGenerateColumns="False" AllowPaging="True" PageSize="15" 
                    CssClass="ListTable" OnPageIndexChanging="GridView1_PageIndexChanging" OnRowDataBound="GridView1_RowDataBound" >
        <Columns>
          <asp:BoundField DataField="ctrootid" HeaderText="ID" Visible="False" />
          <asp:TemplateField HeaderText="主題館名稱">
            <ItemTemplate>
              <asp:LinkButton ID="Namelink" runat="server" Text='<%# Bind("ctrootname") %>'></asp:LinkButton>
            </ItemTemplate>
          </asp:TemplateField>
		 <asp:TemplateField HeaderText="復原主題館">
            <ItemTemplate>
			<asp:Button ID="restore" runat="server" Text="復原" class="cbutton"/> 
			</ItemTemplate>
          </asp:TemplateField>
        </Columns>
        <PagerSettings Visible="False" />
      </asp:GridView>
    
      <asp:SqlDataSource ID="SqlDataSource1" runat="server"></asp:SqlDataSource>

    </form>
        
  </body>
</html>
