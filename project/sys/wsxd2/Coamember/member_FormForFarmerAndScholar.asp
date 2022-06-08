<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/10 上午 09:59:37
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\sys\wsxd2\Coamember\member_FormForFarmerAndScholar.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("usrid")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array()
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage
ChkNumeric chkNumericArray, chkNumericMessage
ChkDate chkDateArray, chkDateMessage
'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=request(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=request(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=request(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then response.redirect(onErrorPath)
	response.end
End Sub

'Log 放在網站根目錄下的 /Logs，檔名： YYYYMMDD_log4U.txt
Function Log4U(strLog)
	if (activeLog4U) then
		fldr=Server.mapPath("/") & "/Logs"
		filename=Year(Date()) & Right("0"&Month(Date()), 2) & Right("0"&Day(Date()),2)
		
		filename = filename & "_log4U.txt"
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'產生新的目錄
		If (Not fso.FolderExists(fldr)) Then
			Set f = fso.CreateFolder(fldr)
		Else
			ShowAbsolutePath = fso.GetAbsolutePathName(fldr)
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'開啟檔案
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile( fldr & "\" & filename , ForAppending, True, -1)
		f.Write strLog  & vbCrLf
	end if
End Function
'##### 結束：此段程式碼為自動產生，註解部份請勿刪除 #####
%><!--#Include virtual = "/site/coa/inc/client.inc" -->
<!--#Include virtual = "/site/coa/inc/dbFunc.inc" -->
<!--#Include file = "CheckFunction.inc" -->

<script language="JavaScript">
<!--
function send(){
	var form=document.form1; 
	//if(CheckAccount(form) == false){
	//	form.account.focus();
	//	return false;
	//}
	//else 
	if(form.realname.value==""){
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
	//else if(CheckPasswd(form) == false){
	//	return false;
	//}
	else if(CheckEmail(form) == false){
		form.email.focus();
		return false;
	}
	//else if(CheckIDn(form) == false){
	//	form.idn.focus();
	//	return false;
	//}
	return true;
}  
//-->
</script>


<%
    Set RSreg = Server.CreateObject("ADODB.RecordSet")
    Response.Buffer = true

    sql2="select * from Member Where account ='" & Request("usrid") & "'"
    set rs = conn.Execute(sql2)
%>
<!-- Content Start-->

          <div class="Maintitle">
		      <!-- InstanceBeginEditable name="Topic" -->加入會員<!-- InstanceEndEditable -->
		  </div>
          <div valign="top" >
      <!-- InstanceBeginEditable name="body" -->
			<form name="form1" method="post" action="" class="FormA">
        <h1>請填寫會員資料</h1>
        <table  cellspacing="0">
          <tr>
            <th><label for="account">*會員帳號：</label></th>
            <td><input name="" type="text" class="Text" id="account" size="30">
        限用英文與數字，可用『-』或『_』，30碼以下</td>
          </tr>
          <tr>
            <th><label for="name">*真實姓名：</label></th>
            <td><input name="" type="text" class="Text" id="name" size="30"></td>
          </tr>
         <tr>
            <th><label for="nickname">暱　　稱：</label></th>
            <td><input name="nickname" type="text" class="Text" id="nickname" size="30"></td>
          </tr>
          <tr>
            <th><label for="pw">*設定密碼：</label></th>
            <td><input name="" type="password" class="Text" id="pw" size="30">
        請自訂英文（區分大小寫）、數字，需同時包含至少1英文和1數字、不含空白及@，6碼以上、16碼以下</td>
          </tr>
          <tr>
            <th><label for="pw">*密碼確認：</label></th>
            <td><input name="" type="password" class="Text" id="pw" size="30"></td>
          </tr>
          <tr>
            <th><label for="idn">*身分證字號：</label></th>
            <td><input name="" type="text" class="Text" id="idn" size="30">
        外籍人士請填護照號碼</td>
          </tr>
          <tr>
            <th><label for="class">*身分類別：</label></th>
            <td>
							<input name="" id="class" type="radio" value="radiobutton">
							研究員
							<input name="" id="class" type="radio" value="radiobutton">
							教職人員
							<input name="" id="class" type="radio" value="radiobutton">
							學生
						</td>
          </tr>
          <tr>
            <th><label for="org">*所屬機關名稱：</label></th>
            <td><input name="" id="org" type="password" size="30" class="Text"></td>
          </tr>
          <tr>
            <th><label for="ophone">*所屬機關電話：</label></th>
            <td><input name="" id="ophone" type="password" size="30" class="Text">
						（公）
							<label for="ext">分機：</label>
							<input name="" id="ext" type="text" size="5" class="Text">
						範例：02-25076627 </td>
          </tr>
          <tr>
            <th><label for="ptitle">*職稱：</label></th>
            <td><input name="" id="ptitle" type="password" size="30"class="Text"></td>
          </tr>
          <tr>
            <th><label for="birthday">出生日期：</label></th>
            <td>西元
                <input name="" type="text" class="Text" id="birthday" size="5">
        年
        <select name="" id="birthday">
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
        <select name="" id="birthday">
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
        日 </td>
          </tr>
          <tr>
            <th><label for="sex">性別：</label></th>
            <td><input name="" type="radio" value="radiobutton" id="sex">
              男
                <input name="" type="radio" value="radiobutton" id="sex">
                女 </td>
          </tr>
          <tr>
            <th><label for="addr">地址：</label></th>
            <td><input name="" type="text" class="Text" id="addr" size="60">
                <label for="zipcode"> 郵遞區號：</label>
                <input name="" type="text" size="5" class="Text" id="zipcode">
                <a href="#">郵遞區號查詢</a> </td>
          </tr>
          <tr>
            <th><label for="phone">電話：</label></th>
            <td>私
                <input name="" type="text" class="Text" id="phone" size="16">
                <label for="ext">分機：</label>
                <input name="" type="text" class="Text" id="ext" size="5">
        範例：02-25076627<br>
        <label for="celphone">手機號碼：</label>
        <input name="" type="text" class="Text" id="celphone" size="16">
        範例：0911123456</td>
          </tr>
          <tr>
            <th><label for="fax">傳真：</label></th>
            <td><input name="" type="text" class="Text" id="fax" size="16">
        範例：02-25076627</td>
          </tr>
          <tr>
            <th><label for="email">E-mail：</label></th>
            <td><input name="" type="text" class="Text" id="email" size="30">
		<input name="checkbox" type="checkbox" value="Y" checked />是否訂閱電子報
              <br>        供系統認證之用，請務必填寫正確</td>
          </tr>
          <tr>
            <th><label for="">自我介紹：</label></th>
            <td><textarea name="" id="" cols="60" rows="5" wrap="VIRTUAL">您可於欄位中輸入簡短的個人簡介</textarea></td>
          </tr>
          <tr>
            <th><label for="photo">上傳個人照片：</label></th>
            <td><input name="" id="" type="file" class="Text" size="40">
                <br>
    您可以上傳您個人電腦上的照片檔案，請按〔瀏覽〕選擇要上傳的檔案。 </td>
          </tr>
          <tr>
            <th><label for="">研究領域及專長：</label></th>
            <td>&nbsp;</td>
          </tr>
        </table>
        <table cellspacing="0" class="DataTb">
          <caption>
  您的產銷資料如下
          </caption>
          <tr>
            <th>產銷班別：</th>
            <td>陽明山產銷18班</td>
          </tr>
          <tr>
            <th>產地（縣市）：</th>
            <td>台北</td>
          </tr>
          <tr>
            <th>產地（鄉鎮）：</th>
            <td>陽明山</td>
          </tr>
        </table>
        <input name="Submit" type="submit" class="Button" value="確定送出">
        <input name="Reset" type="reset" class="Button" value="重新填寫">
      </form>
			<!-- InstanceEndEditable --> </div>


