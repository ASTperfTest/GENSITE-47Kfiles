<%@ Page Language="VB" MasterPageFile="~/Planting/PlantingMasterPage.master" AutoEventWireup="false" CodeFile="PlantingActFinalRank.aspx.vb" Inherits="Planting_PlantingActRank" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>

	<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Planting/PlantingActIntro.aspx">「虛擬種植，知識升值」－豆仔花園讓你變身種植達人－</a></div>
	
  <ul class="group">
    <li><a href="/Planting/PlantingActIntro.aspx">活動簡介</a></li>
    <li ><a href="/Planting/PlantingActStatement.aspx">活動說明</a></li>
    <li><a href="/Planting/PlantingActPrizes.aspx">抽獎條件與獎項</a></li>
    <li><a href="/Planting/PlantingActRules.aspx">活動辦法</a></li>
     <li ><a href="/Planting/PlantingActRank.aspx">排行榜</a></li>
    <li class="activity"><a href="/Planting/PlantingActFinalRank.aspx">11/26排行榜</a></li>
  </ul>
		
	<div class="pantology2">
    <div class="head"><img src="/xslGip/style1/images/title_05.gif" alt="活動排行榜" /></div>
    <div class="body">
	  <h2>2008/11/26中午12點排行榜</h2>
	  <p>2008/11/26中午12點止活動最後排行榜，待活動後續領獎事宜結束後，再予以更新。</p><br/>	 
     
      <div class="Page">
        第<asp:Label ID="PageNumberText" runat="server" CssClass="Number" />/<asp:Label ID="TotalPageText" runat="server"  CssClass="Number"/>頁，
   	    共<asp:Label ID="TotalRecordText" runat="server" CssClass="Number" />筆資料，     
        <asp:HyperLink ID="PreviousLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_left.gif" ID="PreviousImg" AlternateText="上一頁" />
          <asp:Label ID="PreviousText" runat="server" >上一頁 &nbsp;</asp:Label>
        </asp:HyperLink>
        跳至第<asp:DropDownList ID="PageNumberDDL" AutoPostBack="true" runat="server" /> 頁 &nbsp;
        <asp:HyperLink ID="NextLink" runat="server">
          <asp:image runat="server" ImageUrl="/xslGip/style1/images3/arrow_right.gif" ID="NextImg" AlternateText="下一頁"/>
          <asp:Label ID="NextText" runat="server">下一頁 &nbsp;</asp:Label>
        </asp:HyperLink>，每頁                      
        <asp:DropDownList ID="PageSizeDDL" AutoPostBack="true" runat="server">
          <asp:ListItem Selected="True" Value="15">15</asp:ListItem>                    
          <asp:ListItem Value="30">30</asp:ListItem>
          <asp:ListItem Value="50">50</asp:ListItem>
        </asp:DropDownList>筆資料
      </div>

      <asp:Label ID="TableText" runat="server" Text="" />   
            
      <hr />    
      <div class="top">
        <a href="#" title="top">top</a>
      </div>   
    </div>
    <div class="foot"></div>
	</div>	
</asp:Content>

