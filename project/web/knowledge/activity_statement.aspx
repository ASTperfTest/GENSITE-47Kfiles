<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="activity_statement.aspx.vb" Inherits="activity_statement" title="農業知識入口網-農業知識家" ValidateRequest="false" %>

<%@ Register Src="LeftMenu.ascx" TagName="LeftMenu" TagPrefix="uc1" %>
<%@ Register Src="TabText.ascx" TagName="TabText" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">

  <link href="include/style.css" rel="stylesheet" type="text/css" />

  <uc1:LeftMenu ID="LeftMenu1" runat="server" />
        
  <!--中間內容區--> 
  <td class="center">
  
    <uc2:TabText ID="TabText1" runat="server" />

	  <div class="path">目前位置：<a href="/knowledge/knowledge.aspx">首頁</a>
	  &gt;<a href="/knowledge/activity_statement.aspx">活動說明</a>
	  </div>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td class="menuA"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_introduce.aspx"><img src="image/mn_1.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_rule.aspx"><img src="image/mn_2.gif" width="78" height="28" /></a></td>
        <td width="80" align="left" valign="bottom"><a href="/knowledge/activity_statement.aspx"><img src="image/mn_3h.gif" width="78" height="28" /></a></td>
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
        <img src="image/h1_03.gif" width="120" height="31" />
        <p>本次活動需登入會員方可參加。參加活動之會員需於活動截止日前達成活動之抽獎積分始能參加抽獎。本活動限特定題目之討論才可累積積分，發問題目亦須被選為活動題目才能獲取積分。每日活動問題均有變動，請參加者密切注意。</p>
        發表問題前，建議先搜尋是否已有類似或相同之題目發表，以避免被管理者刪除。積分計算規則如下：<a href="/knowledge/myknowledge_record.aspx" class="link">查詢我的積分</a>
<ol>
          <li>『知識問答你˙我˙他』活動進行期間，原有知識家的互動問答仍照常運作，若無被挑選至活動題目，僅不計入積分。 </li>
          <li>活動積分僅計算被勾選為活動問答題目，以及該題目之相關討論，意見不予計分。 </li>
          <li>活動期間歡迎發問，題目一旦被選為活動題目即列入本次活動討論的範圍，該發問與其下討論者皆可獲得積分。若無被勾選之題目及其相關討論則不計分。 </li>
          <li>若活動題目於活動期間之前即已發表，則活動前發表的討論歸類為『舊有討論』，舊討論之發表者可得基本分數，題目發表者得原有分數。 </li>
          <li>若被選定之題目由『主題館』反饋，則該題目發表者可得加權積分。 </li>
          <li>活動期間，會員於知識家發表之問題，經判斷確為與主題館農作物／農產品相關問題，被反饋回主題館後，該題目發問會員可另得額外積分。 </li>
          <li>單一題目可無限次討論，惟積分計算上限為4分。 <br />
          (同一則題目單人發表討論超過兩次後，後續該人於此項目再討論則不予計分)</li>
  <li>農委會保留關閉問題之權利；活動問題下架後，該題目之『參與討論』功能隨之關閉，原發問與討論仍計分，惟於活動期間內不再開放後續討論。 </li>
  <li>活動問題下架同時，該問題之討論草稿即直接發佈並以一般討論計分，惟農委會保留該發布討論之審核與刪除的權利。 </li>
  <li>在活動期間，農委會保留審核、刪除不適當留言、發問與討論（活動題目之相關討論若非關於該主題知識，亦於此範圍內）的權利。若參加者因不恰當的言論被刪除而影響積分，不得有異議。 </li>
  <li>活動結束後至抽獎前，農委會保留審核不適當之討論（活動題目之相關討論若非關於該主題知識，亦於此範圍內）的權利。若因刪除或不公開該討論而影響活動積分，參賽者不得有異議。 </li>
  <li>重複、抄襲其他討論者的言論，該項內容予以刪除，並視情節重大者處以停權半年為懲，並不得參與本次活動之抽獎。 </li>
  <li>被刪除之討論活動積分會自動扣除。 </li>
  <li><strong>積分達</strong><strong>10</strong><strong>分以上者，即可參加抽獎</strong><strong>; </strong><strong>積分達</strong><strong>20</strong><strong>分以上者，可有兩次抽獎機會，以此類推，積分達</strong><strong>50</strong><strong>分以上者，可有五次抽獎機會。</strong><strong> </strong></li>
  <li>以上抽獎機會，五次為上限，僅為增加中獎機率，無法保證中獎。 </li>
  <li>為求比賽公平，同一得獎者不拘獎項，限得獎一次，若有重複中獎情況，主辦單位得以保留較大獎項給得獎者，並取消較小獎項之得獎資格，以順位候補。 
    <table width="100%" border="0" cellpadding="0" cellspacing="0" class="content_table">
    <tr>
      <th width="15%">項目</th>
      <th width="10%">給分</th>
      <th width="75%">條件說明</th>
    </tr>
    <tr>
      <td align="center">主題館回饋發問</td>
      <td align="center">3分</td>
      <td>由主題館回饋發問，被勾選為活動問題者得3分<br />
        每名會員發問不限次數與最高得分。<br />
        關閉該問題時（停止討論），該計分保留。<br /></td>
    </tr>
    <tr>
      <td align="center">一般發問</td>
      <td align="center">2分</td>
      <td>被勾選為活動問題者得2分（無論該問題發表時間為何）<br />
        每名會員發問不限次數與最高得分<br />
        關閉該問題時（停止討論），該計分保留</td>
    </tr>
    <tr>
      <td align="center">一般討論</td>
      <td align="center">2分</td>
      <td>於活動期間參與活動問題進行討論，每一次討論可得2分。<br />
        每名會員、每題發問之討論計分最多為兩次(4分)，超過的部分不予計分<br />
        當會員之討論被刪除則積分會自動扣除<br />
        關閉該問題時（停止討論），該計分保留<br /></td>
    </tr>
    <tr>
      <td align="center">舊有討論</td>
      <td align="center">1分</td>
      <td>被選為活動問題該項目下的舊有討論（發表於活動開始之前），其討論亦可得基本分數。<br />
        關閉該問題時（停止討論），該計分保留</td>
    </tr>
    <tr>
      <td align="center">額外積分</td>
      <td align="center">1分</td>
      <td>會員於知識家發表之問題，經判斷確為與主題館農作物／農產品相關問題，被反饋回主題館後即可另得額外積分        <br /></td>
    </tr>
  </table>
  </li>
</ol>
      </div></td>
  </tr>
</table>
	  <div class="gotop"><a href="#top"><img src="/xslgip/style3/images/gotop.gif" alt="回到頁面最上方" /></a></div><!--20080829 chiung -->
  </td> 
</asp:Content>

