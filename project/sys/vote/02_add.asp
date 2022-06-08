<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	'HTProgCap = "功能管理"
	'HTProgFunc = "問卷調查"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%	
	
	today_str = date
	today_y = Year(today_str)
	today_m = Month(today_str)
	today_d = Day(today_str)
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function CheckValue() {

	var	fm = document.form1;
	var	v1 = fm.notetype[0].checked;
	var	v2 = fm.notetype[1].checked;
	var	v3 = fm.notetype[2].checked;
	var	v4 = fm.notetype[3].checked;
	var	v5 = fm.notetype[4].checked;
	var	v6 = fm.notetype[5].checked;
	var	v7 = fm.notetype[6].checked;
	var	v8 = fm.notetype[7].checked;
	var	v9 = fm.notetype[8].checked;
	var	bdate, bdy = fm.bdate_y.value, bdm = fm.bdate_m.value, bdd = fm.bdate_d.value;
	var	edate, edy = fm.edate_y.value, edm = fm.edate_m.value, edd = fm.edate_d.value;
	
	
	if ( bdm < 10 )		bdm = '0' + bdm;
	if ( bdd < 10 )		bdd = '0' + bdd;
	if ( edm < 10 )		edm = '0' + edm;
	if ( edd < 10 )		edd = '0' + edd;
	
	bdate = bdy.toString() + bdm.toString() + bdd.toString();
	edate = edy.toString() + edm.toString() + edd.toString();

	
	if ( fm.subject.value == '' ) {
		alert('請輸入主題名稱！');
		fm.subject.focus();
		return false;
	}
	
	if ( bdate > edate ) {
		alert('《起訖日期》設定不正確。起日必須在訖日前！');
		return false;
	}

	if ( fm.haveprize[0].checked && !(v1 || v2 || v3 || v4 || v5 || v6 || v7 || v8 || v9) ) {
		alert('請勾選《記錄問卷填寫人資料》！');
		return false;
	}
	if ( fm.onlyonce[0].checked && !v2 ) {
		alert('若答題《限填一次》，請於《記錄問卷填寫人資料》，並核取《電子郵件信箱》欄位！');
		return false;
	}	


	
	return true;
}

function CheckOnline() {
    if (document.form1.online[0].checked) {
        document.form1.kmonline[0].checked = false; 
        document.form1.kmonline[1].checked = true;
    }
}
function CheckkmOnline() {
    if (document.form1.kmonline[0].checked) {
        document.form1.online[0].checked = false;
        document.form1.online[1].checked = true;
    }
}
//-->
</script>
</head>

<body bgcolor="#ffffff">
問卷調查管理
<table width="100%" height="95%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td>&nbsp;</td>
    <td colspan="3" align="center" valign="top"> 
      <p><br><span class="font2">新增問卷主題</span></p>
      <table border="1" cellspacing="0" cellpadding="4" width="100%" bordercolordark="#FFFFFF">
        <form method="post" name="form1" action="02_add_act.asp" onSubmit="javascript:return CheckValue();">
        <tr> 
          <td colspan="2">
            <p><font color="#993300">使用說明：</font></p>
            <ul>
              <li><font color="#993300">主題名稱：請輸入調查主題名稱。</font></li>
              <li><font color="#993300">題目數量：請設定調查題目數量，最多為30題。</font></li>
            </ul>
          </td>
        </tr>
        <tr> 
          <td>主題名稱：</td>
          <td><input type="text" name="subject" maxlength="100"></td>
        </tr>
        <tr> 
          <td>起訖日期：</td>
          <td> 
            西元 
            <select name="bdate_y">
<%
	for i = 2002 to 2012
		if CInt(i) = CInt(today_y) then
			response.write "<option value='" & i & "' selected>" & i & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & "</option>"		
		end if
	next
%>
            </select>
            年 
            <select name="bdate_m">
<%
	for i = 1 to 12
		if CInt(i) = CInt(today_m) then
			response.write "<option value='" & i & "' selected>" & i & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & "</option>"		
		end if
	next
%>
            </select>
            月 
            <select name="bdate_d">
<%
	for i = 1 to 31
		if CInt(i) = CInt(today_d) then
			response.write "<option value='" & i & "' selected>" & i & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & "</option>"		
		end if
	next
%>
            </select>
            日 到 西元 
            <select name="edate_y">
<%
	for i = 2002 to 2012
		if CInt(i) = CInt(today_y) then
			response.write "<option value='" & i & "' selected>" & i & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & "</option>"		
		end if
	next
%>
            </select>
            年 
            <select name="edate_m">
<%
	for i = 1 to 12
		if CInt(i) = CInt(today_m) then
			response.write "<option value='" & i & "' selected>" & i & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & "</option>"		
		end if
	next
%>
            </select>
            月 
            <select name="edate_d">
<%
	for i = 1 to 31
		if CInt(i) = CInt(today_d) then
			response.write "<option value='" & i & "' selected>" & i & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & "</option>"		
		end if
	next
%>
            </select> 
          </td>
        </tr>
        <tr> 
          <td>是否上線(入口網)：</td>
          <td>
            <input type="radio" name="online" value="1" onclick="CheckOnline()"> yes&nbsp;
            <input type="radio" name="online" value="0" checked> no
          </td>
        </tr>
        <tr> 
          <td>是否上線(加值網)：</td>
          <td>
            <input type="radio" name="kmonline" value="1" onclick="CheckkmOnline()"> yes&nbsp;
            <input type="radio" name="kmonline" value="0" checked> no
          </td>
        </tr>
        <tr> 
          <td valign="top">問卷說明：</td>
          <td><textarea name="questionnote" cols="50" rows="10"></textarea></td>
        </tr>
        <tr> 
          <td><p>題目數量：</p></td>
          <td>
            <select name="questno">
<%
	for i = 1 to 30
		response.write "<option value='" & i & "'>" & i & "</option>"
	next
%>
            </select> 
          </td>
        </tr>
        <tr> 
          <td>答題形式：</td>
          <td>
            <input type="radio" name="jumpquestion" value="0" checked> 逐題顯示&nbsp;
            <input type="radio" name="jumpquestion" value="1"> 跳題
          </td>
        </tr>
        <tr> 
          <td valign="top">是否抽獎：</td>
          <td>
            <input type="radio" name="haveprize" value="1"> yes&nbsp;
            <input type="radio" name="haveprize" value="0" checked> no
            (若需要抽獎，請記錄問卷填寫人資料)
          </td>
        </tr>
        <tr> 
          <td valign="top">記錄問卷填寫人資料：<br>(全不勾表示不記錄)</td>
          <td>
            <input type="checkbox" name="notetype" value="1"> 姓名&nbsp;
            <input type="checkbox" name="notetype" value="2"> 電子郵件信箱&nbsp;
            <input type="checkbox" name="notetype" value="3"> 性別&nbsp;
            <input type="checkbox" name="notetype" value="4"> 年齡&nbsp;
            <input type="checkbox" name="notetype" value="5"> 居住縣市<br> 
            <input type="checkbox" name="notetype" value="6"> 家庭成員人數&nbsp;
            <input type="checkbox" name="notetype" value="7"> 收入&nbsp;
            <input type="checkbox" name="notetype" value="8"> 職業&nbsp;
            <input type="checkbox" name="notetype" value="9"> 學歷&nbsp;
          </td>
        </tr>
        <tr> 
          <td>每人是否限填一次：</td>
          <td>
            <input type="radio" name="onlyonce" value="1"> yes&nbsp;
            <input type="radio" name="onlyonce" value="0" checked> no
            (若限填一次，請記錄問卷填寫人電子郵件信箱)
          </td>
        </tr>
        <tr> 
          <td colspan="2">
            <input type="submit" name="submit" value="確定">
            <input type="button" value="取消" onClick="javascript:location.replace('02.asp');">
          </td>
        </tr>
        </form>
      </table>
    </td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
