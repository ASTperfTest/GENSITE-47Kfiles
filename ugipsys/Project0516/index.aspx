<%@ Page Language="C#" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="index" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
  <head runat="server">
    <title>未命名頁面</title>
    <link href="/Project0516/css/list.css" rel="stylesheet" type="text/css"/>
    <link href="./css/layout.css" rel="stylesheet" type="text/css"/>
    <link href="./css/theme.css" rel="stylesheet" type="text/css"/>
    <link rel="stylesheet" type="text/css" href="./css/jquery.autocomplete.css" />
    <script type="text/javascript" src="./js/jquery-1.3.2.min.js"></script>
    <script type="text/javascript" src="./js/jquery.autocomplete.js"></script>
    <script type="text/javascript" language="javascript">
        $(function() {
        $("#txtSubject").autocomplete('AutoComplete.aspx', { delay: 10 });
        });
    </script>
  </head>
  <body>
    <div id="FuncName">
	    <h1>可維護主題館列表</h1>
	    <div id="ClearFloat"></div>
    </div>
    <form id="form1" runat="server">
      查詢主題館：            
      <asp:TextBox ID="txtSubject" runat="server"></asp:TextBox>
      <asp:Button ID="btnSearch" runat="server" Text="查詢" onclick="btnSearch_Click" />      
      
      <asp:ImageButton ID="add" runat="server" PostBackUrl="~/default.aspx" ImageUrl="./images/bt_newsite.gif" CssClass="newsite"/>
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
        ，是否公開：
        <asp:DropDownList ID="IsPublic" runat="server" AutoPostBack="true" OnSelectedIndexChanged="pagesize_SelectedIndexChanged">
          <asp:ListItem Value="" Text="請選擇" Selected="true"/>
          <asp:ListItem Value="Y" Text="Y"/>
          <asp:ListItem Value="N" Text="N"/>
        </asp:DropDownList>&nbsp;&nbsp;
        <asp:Button ID="musicadd" runat="server" Text="背景音樂維護"  PostBackUrl="bgMusic/bgMusicList.aspx" />
      </div>
		<!-- 要增加欄位請注意index是否會影響到隱藏欄位 -->
      <asp:GridView ID="GridView1" runat="server" DataSourceID="SqlDataSource1" AutoGenerateColumns="False" AllowPaging="True" PageSize="15" 
                    CssClass="ListTable" OnPageIndexChanging="GridView1_PageIndexChanging" OnRowDataBound="GridView1_RowDataBound" >
        <Columns>
          <asp:BoundField DataField="ctrootid" HeaderText="ID" Visible="False" />
          <asp:TemplateField HeaderText="主題館名稱">
            <ItemTemplate>
              <asp:LinkButton ID="Namelink" runat="server" Text='<%# Bind("ctrootname") %>'></asp:LinkButton>
            </ItemTemplate>
          </asp:TemplateField>
          <asp:BoundField DataField="inuse" HeaderText="是否公開" Visible="true" />
            <asp:HyperLinkField HeaderText="詞彙維護" NavigateUrl="../Phrase/Phrase.asp" 
                Text="維護" DataNavigateUrlFields="ctrootid,ctrootname,inuse" 
                DataNavigateUrlFormatString="../Phrase/Phrase.asp?mp={0}" 
                DataTextField="ctrootid" DataTextFormatString="維護" />
          <asp:HyperLinkField DataTextField="ctrootid" HeaderText="維護主題文章" NavigateUrl="~/edit/step1_edit.aspx?ID={0}" 
                              DataNavigateUrlFields="ctrootid,ctrootname,inuse" DataNavigateUrlFormatString="~/edit/step1_edit.aspx?id={0}" 
                              DataTextFormatString="維護主題館" Text="維護文章" />
							  <asp:BoundField DataField="ViewCount" HeaderText="瀏覽統計" Visible="true" />
		<asp:BoundField DataField="order_num" HeaderText="排序" Visible="true" />
		  <asp:HyperLinkField DataTextField="UserName" HeaderText="維護專家" NavigateUrl="~/subject_set.asp?ID={0}" 
                              DataNavigateUrlFields="ctrootid,ctrootname,inuse" DataNavigateUrlFormatString="/subjectset/subject_set.asp?id={0}" 
                              DataTextFormatString="{0}" Text="新增" Visible="true"/>
		 <asp:TemplateField HeaderText="封存主題館">
            <ItemTemplate>
			<asp:Button ID="Old" runat="server" Text="封存" class="cbutton"/> 
			</ItemTemplate>
          </asp:TemplateField>
        </Columns>
        <PagerSettings Visible="False" />
      </asp:GridView>
    
      <asp:SqlDataSource ID="SqlDataSource1" runat="server"></asp:SqlDataSource>

    </form>
        
  </body>
</html>
