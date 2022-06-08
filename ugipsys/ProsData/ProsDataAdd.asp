<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head >
    <title>未命名頁面</title>
    <link href="/css/layout.css" rel="stylesheet" type="text/css"/>
    <link href="/css/form.css" rel="stylesheet" type="text/css"/>
</head>
<body>
    <div name="FuncName">
	    <h1>專家資料／新增專家資料</h1><font size="2">【目錄樹節點: 新增專家資料】</font>
	    <div name="Nav"></div>
	    <div name="ClearFloat"></div>
    </div>
    <div name="FormName">		       
       	單元資料維護&nbsp;<font size="2">【編輯(<font color="red">一般資料式</font>)--主題單元:新增專家資料 / 單元資料:純網頁】</font>
    </div>
    <form name="form1" >
        <div>    
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
           
           <script language="JavaScript">            
            function send(){
	            var form=document.Form1; 
	            
	            if(form.accounttext.value==""){
		            alert("您忘了填寫帳號了！"); 
		            form.accounttext.focus();
		            event.returnValue = false;
		            return;		            
	            }	
	            else{
		            if (form.accounttext.value.length > 30){
			            alert("您所填寫的帳號超過30碼！"); 
			            form.accounttext.focus();
			            event.returnValue = false;
			            return;
		            }
		            if (form.accounttext.value.length < 6){
			            alert("您所填寫的帳號少於6碼！"); 
			            form.accounttext.focus();
			            event.returnValue = false;
			            return;
		            }
		            var i;
		            var ch;
		            for (i=0;i< form.accounttext.value.length;i++){
			            ch = form.accounttext.value.charAt(i);
			            if(ch >= 'a' && ch <= 'z'){
				            //return true;
			            }
			            else if(ch >= 'A' && ch <= 'Z'){
				            //return true;
			            }
			            else if(ch >= '0' && ch <= '9'){
				            //return true;
			            }
			            else if(ch == '-' || ch == '_'){
				            //return true;
			            }
			            else{
				            alert("帳號限用英文與數字，可用『-』或『_』！"); 
				            form.accounttext.focus();
				            event.returnValue = false;
				            return;
		    	        }
		            }
	            }	            
	            if(form.realnametext.value==""){
		            alert("您忘了填寫真實姓名了！"); 
		            form.realnametext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.password1text.value==""){
		            alert("您忘了填寫密碼了！"); 
		            form.password1text.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.password1text.value!=""){
            		
		            if (form.password1text.value.length > 16){
			            alert("您所填寫的密碼超過16碼！"); 
			            form.password1text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else if(form.password1text.value.length < 6){
			            alert("您所填寫的密碼少於6碼！"); 
			            form.password1text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else if(form.password2text.value==""){
			            alert("您忘了填寫密碼確認了！"); 
			            form.password2text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else if(form.password2text.value != form.password1text.value){
			            alert("密碼與密碼確認不符！");
			            form.password2text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else{
			            var i;
			            var ch;
			            var digits = "0123456789";
			            var digitflag = false;
			            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
			            var charflag = false;
			            for (i=0;i< form.password1text.value.length;i++){		
				            ch = form.password1text.value.charAt(i);		
				            if(ch == '\"' || ch == ' ' || ch == '\'' || ch== '\&'){
					            alert("密碼請勿包含『\"』、『'』、『&』或空白"); 
					            form.password1text.focus();
					            event.returnValue = false;
					            return;
				            }
				            if( digits.indexOf(ch) >= 0 ) {
					            digitflag = true;
				            }					
				            if( chars.indexOf(ch) >= 0 ) {
					            charflag = true;
				            }
			            }
			            if( !digitflag ) {
					            alert("密碼請至少包含一數字"); 
					            form.password1text.focus();
					            event.returnValue = false;
					            return;
			            }
			            if( !charflag ) {
					            alert("密碼請至少包含一英文字"); 
					            form.password1text.focus();
					            event.returnValue = false;
					            return;
			            }
		            }		
	            }	     
	            if(form.namentext.value==""){
		            alert("您忘了填寫身份證字號了！"); 
		            form.namentext.focus(); 
		            event.returnValue = false;
		            return;
	            }       
	            if(form.member_orgtext.value==""){
		            alert("您忘了填寫所屬機關名稱了！"); 
		            form.member_orgtext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.com_teltext.value==""){
		            alert("您忘了填寫所屬機關電話了！"); 
		            form.com_teltext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.ptitletext.value==""){
		            alert("您忘了填寫職稱了！"); 
		            form.ptitletext.focus(); 
		            event.returnValue = false;
		            return;
	            }	            
	            if(form.emailtext.value==""){
		            alert("您忘了填寫電子郵件了！"); 
		            form.emailtext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.emailtext.value.indexOf("@")<=-1){
		            alert("您所填寫的電子郵件格式有誤！"); 
		            form.emailtext.focus(); 
		            event.returnValue = false;
		            return;
	            }
            }              
            </script>

        </div>
    </form>
</body>
</html>
