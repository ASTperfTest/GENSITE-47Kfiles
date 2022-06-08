<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="activity_introduce.aspx.vb" Inherits="activity_introduce" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <link href="include/style.css" rel="stylesheet" type="text/css" />

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>
	  &gt;<a href="/knowledge/activity_introduce.aspx">活動簡介</a>	  </div>
	  
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	  <td class="menuA"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_introduce.aspx"><img src="image/mn_1h.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_rule.aspx"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_statement.aspx"><img src="image/mn_3.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/knowledge_activityrank.aspx"><img src="image/mn_4.gif" width="78" height="28" /></a></td>
        <td valign="bottom"><asp:Label ID="TabText" runat="server" Text=""></asp:Label></td>
      </tr>
    </table></td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
    <td class="menuA">&nbsp;</td>
  </tr>
  <tr>
    <td class="content"><div class="content_header"><img src="image/header.gif" width="500" height="135" /></div>
      <div class="content_mid">
       <img src="image/h1_01.gif" width="120" height="31" />
	    <div style="width:100%; text-align:right">轉貼到 
        <a href="http://www.facebook.com/sharer.php?u=kmweb.coa.gov.tw/knowledge/activity_introduce.aspx"><img src="image/share_icon_facebook.gif" width="16" height="16" /></a> <a href="http://www.plurk.com/?qualifier=shares&status=http%3A%2F%2Fkmweb.coa.gov.tw%2Fknowledge%2Factivity_introduce.aspx (%E8%BE%B2%E6%A5%AD%E7%9F%A5%E8%AD%98%E5%85%A5%E5%8F%A3%E7%B6%B2%E3%80%8E%E7%9F%A5%E8%AD%98%E5%95%8F%E7%AD%94%E4%BD%A0%E6%88%91%E4%BB%96%E3%80%8F%E6%9C%89%E7%8D%8E%E5%BE%B5%E5%95%8F%E7%AD%94%E6%B4%BB%E5%8B%95%E9%96%8B%E5%A7%8B%E5%9B%89%EF%BC%812010%2F5%2F10~5%2F31%E6%9C%9F%E9%96%93%E5%8A%A0%E5%85%A5%E6%9C%83%E5%93%A1%E3%80%81%E5%8F%83%E5%8A%A0%E4%BA%92%E5%8B%95%E5%95%8F%E7%AD%94%E5%88%86%E4%BA%AB%E7%9F%A5%E8%AD%98%EF%BC%8C%E5%B0%B1%E8%83%BD%E7%8D%B2%E5%BE%97%E7%A9%8D%E5%88%86%E3%80%81%E5%8F%83%E5%8A%A0%E6%8A%BD%E7%8D%8E%EF%BC%8C%E7%A9%8D%E5%88%86%E8%B6%8A%E5%A4%9A%E6%8A%BD%E7%8D%8E%E6%A9%9F%E7%8E%87%E8%B6%8A%E9%AB%98%EF%BC%8C%E6%9C%89%E6%A9%9F%E6%9C%83%E6%8A%BD%E4%B8%ADPSPgo%E3%80%81Ipod%E3%80%81SOGO%E7%A6%AE%E5%8D%B7%E7%AD%8940%E5%80%8B%E5%A4%A7%E7%8D%8E%EF%BC%81)"><img src="image/share_icon_plurk.gif" width="16" height="16" /></a></div>
        <div style="width:100%; text-align:center"><img src="image/page01_ob01.gif" width="428" height="51" /></div>
        <p/>
		<p>於活動期間內，會員所發表農業相關的問題，經挑選為活動題目，則可獲得積分；為鼓勵會員多閱讀主題館文章並從中發問，被選為活動題目若為主題館作物，則可獲得加權積分。 </p>
        <p>除此之外，透過參與活動題目的討論，將寶貴的『農業知識』分享給大家，亦可累積知識能量積分，共有40個獲獎名額，積分越多抽獎機率越高，就有機會將好禮抱回家！</p>
        <p>2010/5/17~6/7期間加入會員、參加互動問答分享知識，還能獲得積分、參加抽獎，這麼難得的機會，千萬不要錯過喔！ </p>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" class="content_table">
          <tr>
            <th width="15%" align="center">獎項</th>
            <th width="70%" align="center">獎品內容</th>
            <th width="15%" align="center">數量</th>
          </tr>
          <tr>
            <td align="center">頭獎</td>
            <td align="center"><a href="http://www.sonystyle.com.tw/is-bin/INTERSHOP.enfinity/WFS/Sony-SonyStyle-Site/zh_TW/-/TWD/ViewProductDetail-Start?productSKU=PSP-N1007PB#productContain" target="_blank">Sony PSP go主機(PSP-N1007鋼琴黑)</a></td>
            <td align="center">1</td>
          </tr>
          <tr>
            <td align="center">貳獎 </td>
            <td align="center"><a href="http://www.apple.com/tw/ipodnano/specs.html" target="_blank">Apple iPod nano 5代 16G(紅、藍、銀)</a></td>
            <td align="center">3</td>
          </tr>
          <tr>
            <td align="center">參獎</td>
            <td align="center">SOGO 3,000元禮卷</td>
            <td align="center">8</td>
          </tr>
          <tr>
            <td align="center">普獎 </td>
            <td align="center">NIO金屬霧面 2.5吋 160GB行動硬碟</td>
            <td align="center">28</td>
          </tr>
        </table>
        <ul>
          <li>活動贈品無法由得獎者指定顏色 </li>
          <li>贈品相關規格請點選圖片查看，詳細介紹請自行洽詢該產品官網 </li>
          <li>主辦單位保留修改活動辦法、活動獎項及審核中獎資格之權利 </li>
        </ul>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" id="product">
          <tr>
            <td align="center" valign="top"><a href="http://www.sonystyle.com.tw/is-bin/INTERSHOP.enfinity/WFS/Sony-SonyStyle-Site/zh_TW/-/TWD/ViewProductDetail-Start?productSKU=PSP-N1007PB#productContain" target="_blank"><img src="image/present_pspgo.jpg" width="108" height="84" /></a><br />
            Sony PSP go主機</td>
            <td align="center" valign="top"><a href="http://www.apple.com/tw/ipodnano/specs.html" target="_blank"><img src="image/present_ipod.jpg" width="108" height="84" /></a><br />
            Apple iPod nano 5代 16G</td>
            <td align="center" valign="top"><img src="image/present_coupon.jpg" width="108" height="84" /><br />
            SOGO 3,000元禮卷</td>
            <td align="center" valign="top"><img src="image/present_hd.jpg" width="108" height="84" /><br />
            NIO金屬霧面 2.5吋 160GB行動硬碟</td>
          </tr>
        </table><p/>
    </div></td>
  </tr>
</table>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td> 
</asp:Content>

