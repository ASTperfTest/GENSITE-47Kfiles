<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="MemberRanking.aspx.vb" Inherits="Member_MemberRanking" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

<a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>

	<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Member/MemberRanking.aspx">會員排行榜</a></div>
	
  <h3>會員排行榜</h3>
		
	<div id="Magazine">
		<div class="Event">
		
		  <div class="browseby">	
        <asp:label ID="browseby" runat="server" text=""></asp:label>
      </div>   
      
      <div class="Page">
        第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
        <asp:HyperLink ID="PreviousLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style3/images/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
          <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
        </asp:HyperLink>
        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
        <asp:HyperLink ID="NextLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style3/images/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
          <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
        </asp:HyperLink>，每頁                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
          <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
          <asp:ListItem Value="30">30</asp:ListItem>
          <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>筆資料
      </div>

      <ul class="subjectcata">
        <asp:label runat="server" text="" ID="TabText"></asp:label>
      </ul>
      <div class="lp">	
        <asp:Label ID="TableText" runat="server" Text="" />   
      </div>  
      <hr />    
      <div class="top">
        <a href="#" title="top">top</a>
      </div>   
    </div>
    <div class="foot"></div>
	</div>	

</asp:Content>

