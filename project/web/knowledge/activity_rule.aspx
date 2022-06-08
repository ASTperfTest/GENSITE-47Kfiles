<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="activity_rule.aspx.vb" Inherits="activity_rule" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <link href="include/style.css" rel="stylesheet" type="text/css" />

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>
	  &gt;<a href="/knowledge/activity_rule.aspx">活動辦法</a>
	  </div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
	<td class="menuA"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_introduce.aspx"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_rule.aspx"><img src="image/mn_2h.gif" width="78" height="28" /></a></td>
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
        <img src="image/h1_02.gif" width="120" height="31" />
        <ol>
          <li>活動時程：2010年5月17日上午10時起~6月7日上午10時止。</li>
          <li>全國民眾凡具中華民國國籍，不限年齡、職業、性別，加入成為本網站會員後，皆可參加此活動。<a href="/sp.asp?xdURL=coamember/member_Join.asp&mp=1" class="link">馬上加入會員 </a></li>
          <li>本次活動新增會員資格認證機制，參賽會員請以中文姓名註冊，並填寫真實有效之個人資料與連絡方式，各會員請勿使用相同連絡信箱或臨時信箱註冊。若您已用英文拼音註冊，請於活動期間內，來信至<a href="mailto:km@mail.coa.gov.tw">km@mail.coa.gov.tw</a>，確認您的身份，請管理員為您修改成中文名字。 </li>
          <li>發表問題前，建議先搜尋是否已有類似或相同之題目發表，以避免被管理者刪除。 </li>
          <li>於推廣活動時間內，玩家發問與討論次數無上限，唯與主題無關或不適當之言論，管理者保留刪除或不公開言論之資格，並不列入活動積分，參加者不得有異議。 </li>
          <li>發問與討論之言論若有抄襲、刻意重覆前發表者之行為嚴重者，將處以會員停權半年的處分，並不得以參與此次活動之抽獎。 </li>
          <li>如有任何因電腦、網路、電話、技術或不可歸責於主辦單位之事由，而使參加者登錄之資料有延遲、遺失、錯誤、無法辨識或毀損之情況，主辦單位不負任何法律責任，參加者亦不得因此異議。 </li>
          <li>為求比賽公平，同一得獎者不拘獎項，限得獎一次，若有重複中獎情況，主辦單位得以保留最大獎項給得獎者，並取消較小獎項之得獎資格，以順位候補。若有得獎者會員之註冊信箱帳號重複，亦視同重複中獎情況處理。得獎者若使用臨時信箱註冊，則視為得獎無效，以順位候補。 </li>
          <li>正式得獎名單應以本網站公佈資料為準。得獎名單將於活動結束後2週內公佈於活動網站，並將以e-mail及電話方式通知得獎者，中獎民眾須於期限內依通知進行身份確認，逾期、或身份不符者，皆視為放棄得獎資格，不另候補。各獎項將於公佈名單後1個月內陸續寄送。 </li>
          <li>得獎者須正確填寫領獎收據，並附上身分證影本以便進行核對領獎身分。得獎者個人資料僅供本次核對領獎身分用，不予退還。 </li>
          <li>活動得獎贈品僅限郵寄台、澎、金、馬地區。 </li>
          <li>贈品寄出將再以e-mail與電話方式通知得獎者，未收到贈品者請於兩個月內提出詢問反應，逾期不受理亦不擔負賠償責任。 </li>
          <li>參賽者若有違反本競賽規則及本網站會員註冊同意書條款所列之規定，一經告發查證，主辦單位可逕行刪除活動資格及會員停權作業；該會員若為得獎者，則獎項將追回，採順位方式候補，不另行通知。相關法律刑責將由參賽者自行負責，若造成第三者之權益損失，參賽者得自負完全法律責任，不得異議。本網站亦將保留相關法律追訴之權利。 </li>
          <li>活動期間為維護比賽公平性，禁止網友從事外掛程式或任何影響活動公平之行為，一旦查證告發，主辦單位有權不需說明，直接刪除其抽獎與得獎資格，並可採順位候補不另行通知。 </li>
          <li>本活動所涉及的個人資料蒐集、運用以及關於個人資料的保密，將適用農委會網站的隱私權保護政策。 </li>
          <li>活動內容如遇不可抗拒之因素須進行變更，應以本網站公告為依據。主辦單位保留修改活動辦法、活動獎項及審核中獎資格之權利。
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
              <td align="center"><a href="http://www.apple.com/tw/ipodnano/specs.html">Apple iPod nano 5代 16G(紅、藍、銀)</a></td>
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
              <li>活動贈品無法由得獎者指定顏色              
              <li>贈品相關規格請點選圖片查看，詳細介紹請自行洽詢該產品官網              
              <li>主辦單位保留修改活動辦法、活動獎項及審核中獎資格之權利
                
            </ul>
          </li>
        </ol>
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
</table>
      </td>
  </tr>
</table>

	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td> 
</asp:Content>

