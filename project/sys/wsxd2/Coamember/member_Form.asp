<!--#Include virtual = "/inc/InvitationCode.inc" -->
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include file = "CheckFunction.inc" -->
<script language="JavaScript">
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
	else if(CheckPasswd(form) == false){
		return false;
	}
	else if(CheckEmail(form) == false){
		form.email.focus();
		return false;
	}
//	else if(CheckIDn(form) == false){
//		form.idn.focus();
//		return false;
//	}
	else if(CheckBirthday(form) == false){
		form.birthYear.focus();
		return false;
	}
	return true;
}

function BirthYearKeyPress()
{
	//是否Key in數字
	if(event.keyCode>=48 && event.keyCode<=57)
		return true;	
	else
		return false;
}  
</script>

	<div class="path" >目前位置：
		<a title="首頁" href="mp.asp">首頁</a>
		&gt;<a href="sp.asp?xdURL=coamember/member_Join.asp" title="加入會員">加入會員</a>
		&gt;<a href="sp.asp?xdURL=coamember/member_Agree.asp&checkboxC=1" title="同意聲明">同意聲明</a>
		&gt; 填寫基本資料
		</div>
		<h3>加入會員</h3>
	<div id="Magazine">
						<div class="Event">	

	<div class="experts">
			
<form name="form1" method="post"  class="FormA" action="sp.asp?xdURL=coamember/member_Applydeal.asp" onSubmit="return send()">			
								
				<p>請填寫會員資料</p>
			  <table  cellspacing="0" class="DataTb1">
          <tr>
            <th><pre><label for="account">*會員帳號：</label></pre></th>
            <td><input name="account" type="text" class="Text" id="account" size="30" maxLength="30">
                <input type="hidden" name="invitationCode" id=invitationCode value=<%=getInvitationCode() %> />
                <input type="hidden" name="invitationCodeType" id=invitationCodeType value=<%=getInvitationCodeType() %> />
              <br/>限用英文或數字，可用『-』或『_』，30碼以下 (建議使用英文加數字以提升帳號安全性)</td>
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
            <td><input name="passwd" type="password" class="Text" id="passwd" size="30" maxLength="16">
              <br/>請自訂英文（區分大小寫）、數字，需同時包含至少1英文和1數字、不含空白及@，6碼以上、16碼以下</td>
          </tr>
          <tr>
            <th><label for="passwd2">*密碼確認：</label></th>
            <td><input name="passwd2" type="password" class="Text" id="passwd2" size="30" maxLength="16"></td>
          </tr>
          <!--<tr>
            <th><pre><label for="idn">*身分證字號：</label></pre></th>
            <td><input name="idn" type="text" class="Text" id="idn" size="30">
              <br/>外籍人士請填護照號碼</td>
          </tr>-->
		  <tr>
			<th><pre><label for="sunRegion">*日出日落顯示：</label></pre></th>
			<td><select name="sunRegion" id="sunRegion">
				<%'日出日落顯示區域，從資料庫中取得
				Dim region
					sqlRegion = "select distinct isnull(xKeyword,'台北') as region from dbo.CuDTGeneric where ictunit=303"
				Set regionSorce = conn.Execute(sqlRegion)
				do while not regionSorce.eof
					region = trim(regionSorce("region"))
				%>
				<option value="<%=region%>" <%if region=sunRegion then response.Write "selected"%>><%=region%></option>
				<% regionSorce.MoveNext
				loop
				%>
			</select></td>
		 </tr>
          <tr>
            <th><label for="birthday">*出生日期：</label></th>
            <td>西元
              <input name="birthYear" type="text" class="Text" id="birthYear" size="5" maxLength="4" onkeypress = "return BirthYearKeyPress()">
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
		    <th>
			  <label for="zip"> 郵遞區號：</label></th>
			  <td>
             <input name="zip" type="text" size="5" class="Text" id="zip"></td>
             <!--a href="#">郵遞區號查詢</a-->
			
		  </tr>
		    
          <tr>
            <th><label for="addr">地址：</label></th>	
            <td>			
             <input name="homeaddr" type="text" class="Text" id="homeaddr" size="36"></td>
              
			
			
          </tr>
                    <tr>
            <th><label for="phone">電話：</label></th>
            <td>
              <input name="phone" type="text" class="Text" id="phone" size="16">
              <label for="home_ext">分機：</label>
              <input name="home_ext" type="text" class="Text" id="home_ext" size="5">
              範例：02-25076627
              </td>
          </tr>
          <tr>
              <th><label for="mobile">手機號碼：</label></th>
            <td>
              <input name="mobile" type="text" class="Text" id="mobile" size="16">
              範例：0911123456
            </td>
          </tr>
          <tr>
            <th><label for="fax">傳真：</label></th>
            <td><input name="fax" type="text" class="Text" id="fax" size="16">
              範例：02-25076627</td>
          </tr>
          <tr>
            <th valign="top"><label for="email">*E-mail：</label></th>
            <td>
			<input name="email" type="text" class="Text" id="email" size="30">
			<input name="orderepaper" type="checkbox" value="Y" checked />是否訂閱電子報
              <br/>*供系統認證之用，請務必填寫正確<br/><font color="red">**在您送出資料後之後，系統會發一封Email認證信給您，請您完成Email認證程序。</font>
			</td>
          </tr>
		  <!--2011.06.15 modify by Grace-->
		  <tr>
            <th><label for="cursorcheck">開啟動態游標：</label></th>
            <td colspan=2><input id="cursorcheck" type="checkbox" name="cursorcheck" value="Y" />是否開啟動態游標</td>
          </tr>
        </table>
			  <input name="Submit" type="submit" class="Button" value="確定送出">
			  <input name="Reset" type="reset" class="Button" value="重新填寫">
			</form>
			
</div>
</div>
</div>


