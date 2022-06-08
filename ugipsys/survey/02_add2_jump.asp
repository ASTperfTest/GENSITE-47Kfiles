<%@ CodePage = 65001 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%

subjectid = request("subjectid")
	questno = request("questno")
	jumpquestion = request("jumpquestion")
	if subjectid = "" or questno = "" or jumpquestion = "" then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="css.css">
</head>
<script language="JavaScript" type="text/JavaScript">
<!--
function CheckValue() {
	
	var	i, j, frm = document.form1, checked_flag;
	
	for ( i = 0; i < frm.elements.length; i++ )
		if ( frm.elements[i].type == 'text' && frm.elements[i].open == '' && frm.elements[i].value == '') {
			alert('請輸入所有題目及答案！');
			return false;
		}
	
	/*
	for ( j = 1; j <= <%= questno %>; j++ )	{
		checked_flag = false;
		for ( i = 0; i < frm.elements.length; i++ )
			if ( frm.elements[i].type == 'checkbox' && frm.elements[i].name.substring(0, 9) == 'default' + j + '-' && frm.elements[i].checked )
				checked_flag = true;
				
		if ( !checked_flag ) {
			alert('請設定《題目' + j + '》的答案預設值！');
			return false;
		}
	}
	*/

	return true;  
}

function checkSelect(quest_type, questno, answerno) {

	var	i, fm = document.form1;
	
<% if jumpquestion = "1" then %>

	if ( quest_type.value == '1' )
	
<% else %>

	if ( quest_type[0].checked )
	
<% end if %>
		for ( i = 0; i < fm.elements.length; i++ )
			if ( fm.elements[i].type == 'checkbox' && fm.elements[i].name.substring(0, 9) == 'default' + questno + '-' && fm.elements[i].name != 'default' + questno + '-' + answerno )
				fm.elements[i].checked = false;
}
//-->
</script>
<body bgcolor="#ffffff">
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
    <!--td class="FormLink" valign="top" align=right><a href="http:adm_inves.asp" title="廣告列表">問卷列表</a--> <!--a href="adm_inves.asp" title="新增廣告">檢視投票結果</a--></td>
  </tr>
  <tr>
    <td class="FormLink" valign="top"><strong><!--font color="#990000">新增問卷</font></strong--><p>（步驟三：輸入問題題目與答案）</p> </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>
  </tr>
<table width="100%" height="95%" border="0" cellpadding="0" cellspacing="0">
  <tr> 

    <td align="center" valign="top">
      <table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">
      
      <form method="post" name="form1" action="02_add2_jump_act.asp" onSubmit="javascript:return CheckValue();">
      <input type="hidden" name="subjectid" value="<% =subjectid %>">
      <input type="hidden" name="questno" value="<% =questno %>">
      <input type="hidden" name="jumpquestion" value="<% =jumpquestion %>">
      <input type="hidden" name="functype" value="<%= functype %>">
        <tr> 
          <td colspan="2"><a name=note></a>
            <p><font color="#993300">使用說明：</font></p>
            <ul>
              <font color="#993300">
              <li>請逐一輸入每一題目的問題，以及答案選項。</li>
              <li>按“重新調查此份問卷”，將進行刪除原作答資料紀錄重新調查，在儲存修改資料後，回列表。</li>
              <li>按“上一步”，回上一個修改頁面</li>
              <li>按“修改資料”，將進行單一題目修改功能，不刪除原作答紀錄，在儲存資料後，回列表。</li>
              </font>
            </ul>
          </td>
        </tr>

<%
	sql = "" & _
		" select m012_title, m012_type from m012 where m012_subjectid = " & subjectid & _
		" and m012_questionid <= " & questno & " order by m012_questionid "
	set rs = conn.execute(sql)
	
	if rs.EOF then
		op_mode = "insert"
	end if
	
	sql = "" & _
		" select '*' + ltrim(str(m015_questionid)) + '*' + ltrim(str(m015_answerid)) + '*' " & _
		" from m015 where m015_subjectid = " & subjectid
	set rs3 = conn.execute(sql)
	while not rs3.EOF
		open_str = open_str & rs3(0) & ","
		rs3.MoveNext
	wend

	for i = 1 to questno		
		title = ""
		m012_type = "1"
		if not rs.EOF then
			title = replace(trim(rs("m012_title")), chr(34), "&quot;")
			m012_type = trim(rs("m012_type"))
			rs.MoveNext
		end if
		if m012_type = "" then
			m012_type = "1"
		end if
		
		if jumpquestion <> "1" then
%>        
        <tr>
          <td rowspan="2">
            題目<% =i %>：
<%
			'if op_mode <> "insert" then
%>
            <!--
            <br>
            <input type="submit" name="submit" value="修改題目<% =i %>">
            -->
<%
			'end if
%>
          </td>
          <td>
            本題答題方式： 
            <input type="radio" name="type<% =i %>" value="1"<% if m012_type = "1" then %> checked<% end if %>> 單選&nbsp;
            <input type="radio" name="type<% =i %>" value="2"<% if m012_type = "2" then %> checked<% end if %>> 複選
          </td>
        </tr>
<%
		end if
%>
        <tr>
<%
		if jumpquestion = "1" then
%>
	  <input type="hidden" name="type<% =i %>" value="1">
	  <td>
	    題目<% =i %>：
<%
			'if op_mode <> "insert" then
%>
	    <!--
	    <br>
	    <input type="submit" name="submit" value="修改題目<% =i %>">
	    -->
<%
			'end if
%>
	  </td>
<%
		end if
%>
          <td>
            <b>題目<% =i %>：</b><input type="text" name="title<% =i %>" value="<% =title %>"><br>
<%
		answerno = request("answerno" & i)
%>
	    <input type="hidden" name="answerno<% =i %>" value="<% =answerno %>">
<%
		sql = "" & _
			" select m013_title, m013_default, m013_nextord from m013 " & _
			" where m013_subjectid = " & subjectid & " and m013_questionid = " & i & _
			" and m013_answerid <= " & answerno & " order by m013_answerid "
		set rs2 = conn.execute(sql)

		for j = 1 to answerno			
			title = ""
			m013_default = "N"
			if not rs2.EOF then
				title = replace(trim(rs2("m013_title")), chr(34), "&quot;")
				m013_default = rs2("m013_default")
				nextord = rs2("m013_nextord")
				rs2.MoveNext
			end if
%>
            答案 <% =j %>：<input type="text" name="answer<% =i %>-<% =j %>" value="<% =title %>">&nbsp;&nbsp;
            
<%
			if jumpquestion = "1" and CInt(i) <> CInt(questno) then
%>
            跳至第 
            <select name="nextord<% =i %>-<% =j %>">
<%
				for k = 1 to questno
					response.write "<option value='" & k & "'"
					if CInt(k) = CInt(nextord) or (nextord = "0" and CInt(k) = CInt(i)+1) then
						response.write " selected"
					end if
					response.write ">" & k & "</option>"
				next
%>
              <option value="0"<% if nextord = "0" then %> selected<% end if %>>問卷結束</option>
            </select> 題&nbsp;&nbsp;
<%
			else
%>
	    <input type="hidden" name="nextord<% =i %>-<% =j %>" value="0">
<%
			end if
%>
            <font color="blue">提供開放式答題 </font><input type="checkbox" name="open<% =i %>-<% =j %>" value="Y"<% if instr(open_str, "*" & i & "*" & j & "*") > 0 then %> checked<% end if %>>　
	    設定為預設值 <input type="checkbox" name="default<% =i %>-<% =j %>" value="Y"<% if m013_default = "Y" then %> checked<% end if %> onClick="javascript:checkSelect(form.type<%= i %>, <%= i %>, <%= j %>);">
            <br>
<%
		next
%>
          </td>
        </tr>
<%
	next
%>

        <tr> 
          <td colspan="2">
<%
	if op_mode = "insert" then
%>
            <input type="submit" name="submit" value="確定" onClick="javascript:return confirm('請確認所有欄位均已輸入完成？');">
<%
	else
%>
	    <input type="submit" name="submit" value="重新調查此份問卷" onClick="javascript:return confirm('確定重新調查此份問卷，並將之前調查資料全部清除？');">
	    <input type="submit" name="submit" value="修改資料" onClick="javascript:return confirm('請確認所有欄位均已輸入完成？');">
<%
	end if
%>
            <input type="button" name="back" value="上一步" onClick="javascript:return history.go(-1);"> <a href=#note>使用說明</a>
          </td>
        </tr>
      </form>
      </table>
    </td>

  </tr>
</table>
</body>
</html>
<%
	set rs = nothing
	set rs2 = nothing
	set rs3 = nothing
%>
