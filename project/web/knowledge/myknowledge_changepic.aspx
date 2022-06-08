<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="myknowledge_changepic.aspx.vb" Inherits="knowledge_myknowledge_changepic" title="Untitled Page" ValidateRequest="false" %>

<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <!--path star (style B)-->
	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>&gt;<a href=""#"">我的個人紀錄</a></div>
	  <!--path end-->

	  <h3>我的個人紀錄</h3>
	
	  <!-- 我的發問TAB LIST／開始 -->
	  <div id="Magazine">
		<div id="MagTabs"><asp:Label ID="TabText" runat="server" Text=""></asp:Label></div>	
		<div class="Event">

		  <!--瀏覽排序/開始 -->
			<h5><asp:label runat="server" ID="memberIdText"></asp:label> 的知識檔案－icon設定</h5>
		  <!--瀏覽排序/結束 -->
  
      <table width="100%" class="myset_data">
      <tr>
        <th>&nbsp;</th>
        <th>入門級會員</th>
        <th>進階級會員</th>
        <th>高手級會員</th>
        <th>達人級會員</th>
      </tr>
      <tr>
        <td nowrap="nowrap"><asp:radiobutton runat="server" ID="picTypeA" GroupName="Type" ></asp:radiobutton>花</td>
        <td><img src="../xslGip/style3/images/seticon1-1.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon1-2.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon1-3.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon1-4.jpg" alt="個人成長圖示" /></td>
      </tr>
      <tr>
        <td nowrap="nowrap"><asp:radiobutton runat="server" ID="picTypeB" GroupName="Type" ></asp:radiobutton>樹</td>
        <td><img src="../xslGip/style3/images/seticon2-1.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon2-2.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon2-3.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon2-4.jpg" alt="個人成長圖示" /></td>
      </tr>
      <tr>
        <td nowrap="nowrap"><asp:radiobutton runat="server" ID="picTypeC" GroupName="Type" ></asp:radiobutton>魚</td>
        <td><img src="../xslGip/style3/images/seticon3-1.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon3-2.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon3-3.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon3-4.jpg" alt="個人成長圖示" /></td>
      </tr>
      <tr>
        <td nowrap="nowrap"><asp:radiobutton runat="server" ID="picTypeD" GroupName="Type" ></asp:radiobutton>雞</td>
        <td><img src="../xslGip/style3/images/seticon4-1.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon4-2.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon4-3.jpg" alt="個人成長圖示" /></td>
        <td><img src="../xslGip/style3/images/seticon4-4.jpg" alt="個人成長圖示" /></td>
      </tr>
      </table>

			<div class="btnCenter">
        <asp:button runat="server" CssClass="btn2" Text="確定" ID="SubmitBtn" />
        <asp:button runat="server" CssClass="btn2" Text="取消" ID="CancelBtn" />	  		
			</div>				
	  </div>
	  </div>
	  <!-- 我的發問TAB LIST／結束 -->
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td>
  
</asp:Content>

