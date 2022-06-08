<%@ CodePage = 65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
today_str = date
	today_y = Year(today_str)
	today_m = Month(today_str)
	today_d = Day(today_str)
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="inc/setstyle.css">
<title>問卷調查</title>
</head>
<script language="JavaScript" type="text/JavaScript">
<!--
function CheckValue() {
	if ( document.form1.subject.value == '' ) {
		alert('請輸入主題名稱！');
		document.form1.subject.focus();
		return false;
	}
	
	return true;
}
//-->
</script>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
<form method="post" name="form1" action="02_add_act.asp" onSubmit="return CheckValue();">
  <tr>
	    <td width="100%" class="FormName" colspan="5">單元資料維護&nbsp;
	    <font size=2>【主題單元: 問卷調查】</td>
  </tr>
  <tr>
    <td width="100%" colspan="5">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right><a href="http://hywade.hyweb.com.tw/GipEdit/ctXMLin.asp?ItemID=6&CtNodeID=61" title="廣告列表">問卷列表</a> <!--a href="adm_inves_new.asp" title="新增廣告">新增問卷</a--></td>
  </tr>
  <tr>
    <td class="FormLink" valign="top"><strong><!--font color="#990000">新增問卷</font--></strong>（步驟一：建立問卷基本資料） </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">
      <table width="100%" border="0" cellpadding="5" cellspacing="1">
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>問卷主題：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <input name="subject" type="text" class="inputbox" size="60">
          </td>
        </tr>
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>起訖日期：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">西元 
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
      <select name="bdate_m" class="px12">
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
      <select name="bdate_d" class="px12">
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
      日　至　西元 
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
      <select name="edate_m" class="px12">
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
      <select name="edate_d" class="px12">
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
      日</td>
        </tr>
        <tr>
          <td align="right" valign="top" nowrap class="eTableLable"><strong>問卷說明：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <textarea name="questionnote" cols="58" rows="5" wrap="VIRTUAL" class="textfield"></textarea>
          </td>
        </tr>
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>題目數量：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <select name="questno">
<%
	for i = 1 to 20
		response.write "<option value='" & i & "'>" & i & "</option>"
	next
%>
            </select> 
      題</td>
        </tr>
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>答題方式：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <input name="jumpquestion" type="radio" value="0" checked>
      一般（ 逐題作答）　
      
<input type="radio" name="jumpquestion" value="1">
      跳題</td>
        </tr>
        <tr>
          <td align="right" valign="top" nowrap class="eTableLable"><strong>填卷者資料：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <table border="0" cellpadding="5" cellspacing="0" class="eTableContent">
              <tr>
                <td valign="top" nowrap>
                  <label>
                  <input name="haveprize" type="radio" value="1" checked>
            需要</label>
                </td>
                <td>請勾選資料項目
                    <table width="100%" border="0" cellpadding="3" cellspacing="3" class="eTableContent">
                      <tr>
                        <td><input type="checkbox" name="notetype" value="1">姓名</td>
                        <td><input type="checkbox" name="notetype" value="2">Email</td>
                        <td><input type="checkbox" name="notetype" value="3">性別 </td>
                        <td><input type="checkbox" name="notetype" value="4">年齡</td>
                      </tr>
                      <tr>
                        <td><input type="checkbox" name="notetype" value="5">居住縣市</td>
                        <td><input type="checkbox" name="notetype" value="6">家庭成員數</td>
                        <td><input type="checkbox" name="notetype" value="7">收入</td>
                        <td><input type="checkbox" name="notetype" value="8">職業</td>
                      </tr>
                    </table>
                </td>
              </tr>
              <tr>
                <td valign="top" nowrap>
                  <label>
                  <input type="radio" name="haveprize" value="0">
            不需要</label>
                </td>
                <td>（請注意，勾選本項則無法分析及抽獎）</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>填卷限制：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <label>
            <input name="onlyonce" type="radio" value="0" checked>
    不限制
    <input type="radio" name="onlyonce" value="1">
    每人限填一次</label>
            <label></label>
          </td>
        </tr>
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>問卷隱藏設定：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <label>
            <input name="askhide" type="radio" value="0" checked>
      不隱藏
      <input type="radio" name="askhide" value="1">
      隱藏</label>
            <label></label>
          </td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td width="60%" height="40" align="right"> <a href="javascript:history.go(-1);">back</a>
              <input name="Submit3" type="submit"  value="下一步，題目設計">
          </td>
        </tr>
      </table>
</td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>    
    </tr>
</form>
</table> 

</body>
</html>                                 


    
     
  
     