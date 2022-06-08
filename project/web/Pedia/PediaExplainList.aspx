<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PediaExplainList.aspx.vb" Inherits="Pedia_PediaExplainList" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>
  
  <div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Pedia/PediaList.aspx">農業知識小百科</a></div>
	
	<!-- 頁面功能 Start -->
	<ul class="Function2">
		<li><asp:HyperLink ID="CanAddLink" runat="server" CssClass="add" Visible="false">我來補充</asp:HyperLink></li>
		<li><a href="javascript:history.back()" class="Back">回上一頁</a></li>
	</ul>
	<!-- 頁面功能 end -->
	
	<!-- 內容區 Start -->
	<div class="pantology">
    <div class="head"></div>
    <div class="body">
	    <h2>農業知識小百科</h2>
	    
	    <asp:Label ID="WordTitle" runat="server" Text="" />
    
       <!--分頁/開始 -->
    
      <div class="Page">
        第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
        <asp:HyperLink ID="PreviousLink" runat="server">
          <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
          <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
        </asp:HyperLink>
        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
        <asp:HyperLink ID="NextLink" runat="server">
          <asp:image runat="server" ImageUrl="../xslgip/style3/images/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
          <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
        </asp:HyperLink>，每頁                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
          <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
          <asp:ListItem Value="30">30</asp:ListItem>
          <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>筆資料
      </div>

      <asp:Label ID="TableText" runat="server" Text="" />    
    </div>
    
    <div class="foot"></div>
    
  </div>         

</asp:Content>

