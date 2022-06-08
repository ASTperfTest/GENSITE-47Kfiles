<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->

<script language="JavaScript">
<!--
function send(){
	var form=document.form1; 
	if(CheckAccount(form) == false){
		form.account.focus();
		return false;
	}
	else if(form.realname.value==""){
		alert("您忘了填寫真實姓名了！"); 
		form.realname.focus(); 
		return false;
	}
	else if(form.actor[0].checked==false && form.actor[1].checked==false && form.actor[2].checked==false ){
		alert("您忘了選身份類別了！"); 
		form.actor[0].focus(); 
		return false;
	}
	else if(form.member_org.value==""){
		alert("您忘了填寫所屬機關名稱了！"); 
		form.member_org.focus(); 
		return false;
	}
	else if(form.com_tel.value==""){
		alert("您忘了填寫所屬機關電話了！"); 
		form.com_tel.focus(); 
		return false;
	}
	else if(form.ptitle.value==""){
		alert("您忘了填寫職稱了！"); 
		form.ptitle.focus(); 
		return false;
	}
	else if(CheckPasswd(form) == false){
		return false;
	}
	else if(CheckEmail(form) == false){
		form.email.focus();
		return false;
	}
	else if(CheckIDn(form) == false){
		form.idn.focus();
		return false;
	}
	return true;
}  
//-->
</script>

<div class="path" >目前位置：
		<a title="首頁" href="mp.asp">首頁</a>
		&gt;<a href="sp.asp?xdURL=coamember/member_Join.asp" title="加入會員">加入會員</a>
		&gt;<a href="sp.asp?xdURL=coamember/member_Agree.asp&checkboxS=3" title="同意聲明">同意聲明</a>
		&gt; 填寫基本資料
		</div>
		<h3>加入會員</h3>
	<div id="Magazine">
						<div class="Event">	

	<div class="experts">
<form name="form1" method="post" class="FormA" action="sp.asp?xdURL=coamember/member_ApplydealForScholar.asp" onsubmit="return send()">


				<p>請填寫學者會員資料，有註明＊為必填欄位</p>
			  <table  cellspacing="0" class="DataTb1">
          <tr>
            <th><label for="account">*會員帳號：</label></th>
            <td><input name="account" type="text" class="Text" id="account" size="30">
              <br/>限用英文與數字，可用『-』或『_』，30碼以下</td>
          </tr>
          <tr>
            <th><label for="realname">*真實姓名：</label></th>
            <td><input name="realname" type="text" class="Text" id="realname" size="30"></td>
          </tr>
          <tr>
            <th><label for="nickname">暱　　稱：</label></th>
            <td><input name="nickname" type="text" class="Text" id="nickname" size="30"></td>
          </tr>
          <tr>
            <th><label for="passwd">*設定密碼：</label></th>
            <td><input name="passwd" type="password" class="Text" id="passwd" size="30">
              <br/>請自訂英文（區分大小寫）、數字，需同時包含至少1英文和1數字、不含空白及@，6碼以上、16碼以下</td>
          </tr>
          <tr>
            <th><label for="passwd">*密碼確認：</label></th>
            <td><input name="passwd2" type="password" class="Text" id="passwd2" size="30"></td>
          </tr>
          <tr>
            <th><label for="idn">*身分證字號：</label></th>
            <td><input name="idn" type="text" class="Text" id="idn" size="30">
              <br/>外籍人士請填護照號碼</td>
          </tr>
          <tr>
            <th><label for="actor">*身分類別：</label></th>
            <td><input name="actor" id="actor" type="radio" value="1">
    研究員
      <input name="actor" id="actor" type="radio" value="2">
    教職人員
    <input name="actor" id="actor" type="radio" value="3">
    學生 </td>
          </tr>
          <tr>
            <th><label for="member_org">*所屬機關名稱：</label></th>
            <td><input name="member_org" id="member_org" type="text" size="30" class="Text"></td>
          </tr>
          <tr>
            <th><label for="com_tel">*所屬機關電話：</label></th>
            <td><input name="com_tel" id="com_tel" type="text" size="30" class="Text">
						（公）<label for="com_ext">分機：</label>
							<input name="com_ext" id="com_ext" type="text" size="5" class="Text">
							
						</td>
          </tr>
          <tr>
            <th><label for="ptitle">*職稱：</label></th>
            <td><input name="ptitle" id="ptitle" type="text" size="30"class="Text"></td>
          </tr>
          <tr>
            <th><label for="birthday">出生日期：</label></th>
            <td>西元
              <input name="birthYear" type="text" class="Text" id="birthYear" size="5">
              年
              <select name="birthMonth" id="birthMonth">
                <option>1</option>
                <option>2</option>
                <option>3</option>
                <option>4</option>
                <option>5</option>
                <option>6</option>
                <option>7</option>
                <option>8</option>
                <option>9</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
              </select>
              月
              <select name="birthday" id="birthday">
                <option>1</option>
                <option>2</option>
                <option>3</option>
                <option>4</option>
                <option>5</option>
                <option>6</option>
                <option>7</option>
                <option>8</option>
                <option>9</option>
                <option>10</option>
                <option>11</option>
                <option>12</option>
                <option>13</option>
                <option>14</option>
                <option>15</option>
                <option>16</option>
                <option>17</option>
                <option>18</option>
                <option>19</option>
                <option>20</option>
                <option>21</option>
                <option>22</option>
                <option>23</option>
                <option>24</option>
                <option>25</option>
                <option>26</option>
                <option>27</option>
                <option>28</option>
                <option>29</option>
                <option>30</option>
                <option>31</option>
              </select>
              日
						</td>
          </tr>
          <tr>
            <th><label for="sex">性別：</label></th>
            <td>
							<input name="sex" type="radio" value="1" id="sex">男
              <input name="sex" type="radio" value="0" id="sex">女
						</td>
          </tr>
          <tr>
            <th><label for="homeaddr">地址：</label></th>
            <td>
						 <input name="homeaddr" type="text" class="Text" id="homeaddr" size="60">
            <label for="zip"> 郵遞區號：</label>
             <input name="zip" type="text" size="5" class="Text" id="zip">
            
						</td>
          </tr>
          <tr>
            <th><label for="phone">電話：</label></th>
            <td>
              <input name="phone" type="text" class="Text" id="phone" size="16">
              <label for="home_ext">分機：</label>
              <input name="home_ext" type="text" class="Text" id="home_ext" size="5">
              
              </td>
          </tr>
          <tr>
              <th><label for="mobile">手機號碼：</label></th>
            <td>
              <input name="mobile" type="text" class="Text" id="mobile" size="16">
             
            </td>
          </tr>
          <tr>
            <th><label for="fax">傳真：</label></th>
            <td><input name="fax" type="text" class="Text" id="fax" size="16">
            </td>
          </tr>
          <tr>
            <th><label for="email">*E-mail：</label></th>
            <td><input name="email" type="text" class="Text" id="email" size="30">
		<input name="orderepaper" type="checkbox" value="Y" checked />是否訂閱電子報
              <br/>供系統認證之用，請務必填寫正確</td>
          </tr>
          <tr>
            <th><label for="htx_KMcat">研究領域及專長：</label></th>
					<td>
						<input class="Button" type="text" size="30" name="htx_KMcat"  id="htx_KMcat" >
						<input class="Text" type="hidden" size="60" name="htx_KMcatID"  id="htx_KMcatID" >
						<input class="Text" type="hidden" size="60" name="htx_KMautoID"  id="htx_KMautoID" >
						<BUTTON id="btn_KMcat" class="Text">外部分類</BUTTON>
					</td>
          </tr>
        </table>
			  <input name="Submit" type="submit" class="Button" value="確定送出">
			  <input name="Reset" type="reset" class="Button" value="重新填寫">
			</form>
		</div>
			</div>
				</div>

	<script language="vbs">                     
		document.domain = "coa.gov.tw"
		sub btn_KMcat_onClick
			window.open "http://kmintra.coa.gov.tw/coa/ekm/manage_doc/report2cat22.jsp?data_base_id=DB020&id_name=htx_KMcatID&autoid_name=htx_KMautoID&nm_name=htx_KMcat&&subNode=1*RB1&display=1*none&form_name=form1&focus=&catidsInput=&anchor=",null,"height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no"
		end sub
	</script>
