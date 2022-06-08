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

	subjectid = request("subjectid")
	if subjectid = "" then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	
	
	sql = "" & _
		" select m011_subject, m011_questno, m011_bdate, m011_edate, " & _
		" isNull(m011_online, '0') m011_online, m011_notetype, " & _
		" m011_haveprize, m011_jumpquestion, m011_onlyonce, m011_questionnote, " & _
		" isNull(m011_km_online, '0') m011_km_online " & _
		" from m011 where m011_subjectid = " & subjectid
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	
	notetype = trim(rs("m011_notetype"))
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function CheckValue3() {

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
		alert('若答題《限填一次》，請《記錄問卷填寫人資料》，並核取《電子郵件信箱》欄位！');
		return false;
	}
	
	return true;
}

function CheckValue2(questno, jumpquestion) {

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
	
	if ( fm.questno.value != questno || !fm.jumpquestion[jumpquestion].checked ) {
		alert('《題目數量》、《答題形式》已經變更，請按“繼續修改問卷內容”！');
		return false;
	}
	
	
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
		alert('請記錄問卷填寫人資料！');
		return false;
	}
	
	if ( fm.onlyonce[0].checked && !v2 ) {
		alert('若答題《限填一次》，請《記錄問卷填寫人資料》，並核取《電子郵件信箱》欄位！');
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
    <td align="center" valign="top">
      <p class="font2">修改問卷</p>
      <table border="1" cellspacing="0" cellpadding="4" width="100%" bordercolordark="#FFFFFF">
      <form method="post" name="form1" action="02_add_act.asp">
      <input type="hidden" name="subjectid" value="<% =subjectid %>">
        <tr> 
          <td>主題名稱：</td>
          <td><input type="text" name="subject" maxlength="100" value="<% =replace(trim(rs("m011_subject")), chr(34), "&quot;") %>"></td>
        </tr>
        <tr> 
          <td>起訖日期：</td>
          <td> 
            西元 
            <select name="bdate_y">
<%
	for i = 2002 to 2012
		if CInt(i) = CInt(Year(rs("m011_bdate"))) then
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
		if CInt(i) = CInt(Month(rs("m011_bdate"))) then
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
		if CInt(i) = CInt(Day(rs("m011_bdate"))) then
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
		if CInt(i) = CInt(Year(rs("m011_edate"))) then
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
		if CInt(i) = CInt(Month(rs("m011_edate"))) then
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
		if CInt(i) = CInt(day(rs("m011_edate"))) then
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
          <td>是否上線：</td>
          <td>
            <input type="radio" name="online" value="1" onclick="CheckOnline()"<% if rs("m011_online") = "1" then %> checked<% end if %>> yes&nbsp;
            <input type="radio" name="online" value="0"<% if rs("m011_online") = "0" then %> checked<% end if %>> no
          </td>
        </tr>
        <tr> 
          <td>是否上線(加值網)：</td>
          <td>
            <input type="radio" name="kmonline" value="1" onclick="CheckkmOnline()"<% if rs("m011_km_online") = "1" then %> checked<% end if %>> yes&nbsp;
            <input type="radio" name="kmonline" value="0"<% if rs("m011_km_online") = "0" then %> checked<% end if %>> no
          </td>
        </tr>
        <tr> 
          <td valign="top">問卷說明：</td>
          <td><textarea name="questionnote" cols="50" rows="10"><% =trim(rs("m011_questionnote")) %></textarea></td>
        </tr>
        <tr> 
          <td><p>題目數量：</p></td>
          <td>
            <select name="questno">
<%
	for i = 1 to 30
		if CInt(i) = CInt(rs("m011_questno")) then
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
          <td>答題形式：</td>
          <td>
            <input type="radio" name="jumpquestion" value="0"<% if rs("m011_jumpquestion") = "0" then %> checked<% end if %>> 逐題顯示&nbsp;
            <input type="radio" name="jumpquestion" value="1"<% if rs("m011_jumpquestion") = "1" then %> checked<% end if %>> 跳題
          </td>
        </tr>
        <tr> 
          <td valign="top">是否抽獎：</td>
          <td>
            <input type="radio" name="haveprize" value="1"<% if rs("m011_haveprize") = "1" then %> checked<% end if %>> yes&nbsp;
            <input type="radio" name="haveprize" value="0"<% if rs("m011_haveprize") = "0" then %> checked<% end if %>> no
            (若需要抽獎，請記錄問卷填寫人資料)
          </td>
        </tr>
        <tr> 
          <td valign="top">記錄問卷填寫人資料：<br>(全不勾表示不記錄)</td>
          <td>
            <input type="checkbox" name="notetype" value="1"<% if mid(notetype, 1, 1) = "1" then %> checked<% end if %>> 姓名&nbsp;
            <input type="checkbox" name="notetype" value="2"<% if mid(notetype, 2, 1) = "1" then %> checked<% end if %>> 電子郵件信箱&nbsp;
            <input type="checkbox" name="notetype" value="3"<% if mid(notetype, 3, 1) = "1" then %> checked<% end if %>> 性別&nbsp;
            <input type="checkbox" name="notetype" value="4"<% if mid(notetype, 4, 1) = "1" then %> checked<% end if %>> 年齡&nbsp;
            <input type="checkbox" name="notetype" value="5"<% if mid(notetype, 5, 1) = "1" then %> checked<% end if %>> 居住縣市<br> 
            <input type="checkbox" name="notetype" value="6"<% if mid(notetype, 6, 1) = "1" then %> checked<% end if %>> 家庭成員人數&nbsp;
            <input type="checkbox" name="notetype" value="7"<% if mid(notetype, 7, 1) = "1" then %> checked<% end if %>> 收入&nbsp;
            <input type="checkbox" name="notetype" value="8"<% if mid(notetype, 8, 1) = "1" then %> checked<% end if %>> 職業&nbsp;
            <input type="checkbox" name="notetype" value="9"<% if mid(notetype, 9, 1) = "1" then %> checked<% end if %>> 學歷&nbsp;
          </td>
        </tr>
        <tr> 
          <td>每人是否限填一次：</td>
          <td>
            <input type="radio" name="onlyonce" value="1"<% if rs("m011_onlyonce") = "1" then %> checked<% end if %>> yes&nbsp;
            <input type="radio" name="onlyonce" value="0"<% if rs("m011_onlyonce") = "0" then %> checked<% end if %>> no
            (若限填一次，請記錄問卷填寫人電子郵件信箱)
          </td>
        </tr>
        <tr> 
          <td colspan="2">
	    <input type="submit" name="submit" value="繼續修改問卷內容" onClick="javascript:return CheckValue3();"> 
	    <input type="submit" name="submit" value="刪除本問卷" onClick="javascript:return confirm('您確定刪除？');">
	    <input type="submit" name="submit" value="修改回列表" onClick="javascript:return CheckValue2(<%= rs("m011_questno") %>, <%= rs("m011_jumpquestion") %>);">
	    <input type="button" value="取消" onClick="javascript:location.replace('02.asp');">
          </td>
        </tr>
      </form> 
    </td>
  </tr>
</table>
<table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">
  <tr> 
    <td> 
      <font color="#993300"><b>修改問卷說明：</b><br>
      <ul>
        <li>按“繼續修改問卷內容”，可再進一步修改問卷資料。</li>
        <li>按“刪除本問卷”，刪除所有和此問卷相關資料。</li>
        <li>按“修改回列表”，儲存修改資料後回列表，但有修改《題目數量》、《答題形式》則必須按“繼續修改問卷內容” 進一步修改問卷內容。</li>
        <li>按“取消”，將不修改任何資料，回列表。</li>
      </ul>
    </td>
  </tr>  
</table>    
</body>
</html>
