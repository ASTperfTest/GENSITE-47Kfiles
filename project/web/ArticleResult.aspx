<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false"
    CodeFile="ArticleResult.aspx.vb" Inherits="ArticleResult" Title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" runat="Server">
    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">
        目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;<a title="主題館" href="/SubjectList.aspx">主題館</a>&gt;<a href="#">文章查詢結果</a></div>
    <h3>
        主題館-文章查詢</h3>
    <div class="Page">
            第<asp:label id="PageNumberText" runat="server" text="" cssclass="Number" />/<asp:label
                id="TotalPageText" runat="server" text="" cssclass="Number" />頁， 共<asp:label id="TotalRecordText"
                    runat="server" text="" cssclass="Number" />筆資料，
            <asp:hyperlink id="PreviousLink" runat="server">
    	  <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" Visible="false" AlternateText="上一頁"></asp:image>        	            
    	  <asp:Label ID="PreviousText" runat="server" Visible="false" Text="Label">上一頁 &nbsp;</asp:Label>            
    	</asp:hyperlink>
            到第
            <asp:dropdownlist id="PageNumberDDL" autopostback="true" runat="server" />
            頁</label>&nbsp;
            <asp:hyperlink id="NextLink" runat="server">
        <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" Visible="false" AlternateText="下一頁"></asp:image>                     	
        <asp:Label ID="NextText" runat="server" Visible="false" Text="Label">下一頁 &nbsp;</asp:Label>
      </asp:hyperlink>
            ， 每頁顯示
            <asp:dropdownlist id="PageSizeDDL" autopostback="true" runat="server">
                <asp:ListItem Selected="True" Value="10">10</asp:ListItem>
                <asp:ListItem Value="20">20</asp:ListItem>
                <asp:ListItem Value="30">30</asp:ListItem>
                <asp:ListItem Value="50">50</asp:ListItem>
            </asp:dropdownlist>
            筆 &nbsp;&nbsp;
            <a title="重新查詢" href="/ArticleSearch.aspx">重新查詢</a>
        </div>
        <br /><br />
        <asp:label id="TableText" runat="server" text=""></asp:label>
</asp:Content>
