<%@ Page Language="VB" MasterPageFile="~/4SideMasterPage.master" AutoEventWireup="false" CodeFile="PlantingActRank.aspx.vb" Inherits="Planting_PlantingActRank" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <a href="#" title="網頁內容資料" class="Accesskey" accesskey="C">:::</a>

	<div class="path">目前位置：<a href="/">首頁</a>&gt;<a href="/Planting/PlantingActIntro.aspx">「虛擬種植，知識升值」－豆仔花園讓你變身種植達人－</a></div>
	
  <ul class="group">
    <li><a href="/Planting/PlantingActIntro.aspx">活動簡介</a></li>
    <li ><a href="/Planting/PlantingActStatement.aspx">活動說明</a></li>
    <li><a href="/Planting/PlantingActPrizes.aspx">抽獎條件與獎項</a></li>
    <li><a href="/Planting/PlantingActRules.aspx">活動辦法</a></li>
    <li class="activity"><a href="/Planting/PlantingActRank.aspx">排行榜</a></li>
  </ul>
		
	<div class="pantology2">
    <div class="head"><img src="/xslGip/style1/images/title_05.gif" alt="活動排行榜" /></div>
    <div class="body">
	  <h2>排行榜</h2>
	  <p>＊	請注意!!活動時間從<strong class="red01">2008/10/27中午12點起-2008/11/25中午12點止</strong>，活動期間以外的得分與星級將不納入抽獎資格計算。</p><br/>
	  <p>＊	活動結束前，即使尚未種植成功，也會將當時種植分數轉換成星級，累加至玩家星級中，以鼓勵玩家持續參與活動至最後。</p><br/>
    <p>＊	本活動排行榜<strong class="red01">每小時</strong>更新一次，於活動結束時間（2008/11/25中午12點）排行榜將暫時凍結，待活動後續領獎事宜結束後，再予以更新。</p>	<br/>
      
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

