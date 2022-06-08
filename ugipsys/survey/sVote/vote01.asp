<%@ CodePage = 65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HEAD>
<TITLE>title</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
</HEAD>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbFunc.inc" -->
<%
	ctNode = replace(trim(request("ctNode")), "'", "''")
	subjectid = replace(trim(request("subjectid")), "'", "''")
	if subjectid = "" then
		response.write "<body onload=JavaScript:alert('操作錯誤！');history.go(-1);>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if

	sql = "select * from m011 where m011_subjectid = " & subjectid
	set rs = conn.execute(sql)

	sql = "select m015_questionid, m015_answerid from m015 where m015_subjectid = " & subjectid
	set rs3 = conn.execute(sql)
	while not rs3.eof
		open_str = open_str & "*" & rs3(0) & "*" & rs3(1) & "*,"
		rs3.movenext
	wend
%>
<script LANGUAGE="JavaScript">
<!--
function send() {
	var form = document.post;
	if ( form.name.value == "" ) {
		alert("請輸入姓名！");
		form.name.focus();
		return false;
	} else if ( form.email.value == "" ) {
		alert("請輸入email！");
		form.email.focus();
		return false;
	}

	return true;
}
//-->
</script>

<!-- 項目 -->
<div class="PathArea"><a href="np.asp?ctNode=464" class="AccessContent" title=":::網站主要內容區" accesskey="C">:::</a><img src="xslgip/forest/images/content/icon_home.gif" alt="回首頁" hspace="3" border="0" align="absmiddle"/><a href="mp.asp" xmlns="">首頁</a><img src="xslgip/forest/images/arrow_path.gif" alt=" " xmlns="" />問卷調查</div> 
<div class="TitleSet">
	<div class="TitleFont"><img src="xslgip/forest/images/icon_tree.gif" alt="標題icon圖" hspace="5" align="absmiddle"/>歡迎使用問卷調查</div>
</div>
<div id="TableSet" xmlns="">
<form method="post" action="sp.asp?xdURL=svote/inquire_qa_act.asp&ctNode=<%= ctNode %>" name="post" onSubmit="javascript:return send();">
<input type="hidden" name="subjectid" value="<%= subjectid %>">
<table cellspacing="0" class="surveyTable" summary="意見調查表格">
  <tr>
    <th colspan="2" width="100%"><font><% =trim(rs("m011_subject"))%></font></th>
  </tr>
        <tr>
          <td colspan="2"><%=trim(rs("m011_body"))%><br></td>
        </tr>
	<tr>
           <td colspan="2" class="IntroContent">請填寫相關資料，以便進行統計分析。</td>
         </tr>
<%
	notetype = trim(rs("m011_notetype"))
	
	if mid(notetype, 1, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="name">姓名：</label></strong></td>
                            <td class="answer"><input type="text" name="m014_A" value="請輸入姓名" tabindex="1" id="name" onFocus="if(this.value=='請輸入姓名')this.value='';"></td>
                          </tr>
<%
	else
		response.write "<input type=""hidden"" name=""name"" value=""none"">"
	end if
	
	if mid(notetype, 2, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="connect">聯絡方式：</label></strong></td>
                            <td class="answer"><input type="text" name="m014_B" value="請輸入聯絡方式" tabindex="2" id="connect" onFocus="if(this.value=='請輸入聯絡方式')this.value='';"></td>
                          </tr>
<% 
	else
		response.write "<input type=""hidden"" name=""email"" value=""@."">"
	end if 
	
	if mid(notetype, 3, 1) = "1" then
%>
                           <tr>
                            <td class="question"><strong><label for="phone">電話：</label></strong></td>
                            <td class="answer"><input type="text" name="m014_C" value="請輸入電話" tabindex="3" id="phone" onFocus="if(this.value=='請輸入電話')this.value='';"></td>
                           </tr>
<% 
	end if 
	
	if mid(notetype, 4, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="email">E-mail：</label></strong></td>
                            <td class="answer"><input type="text" name="m014_D" value="請輸入E-mail" tabindex="4" id="email" onFocus="if(this.value=='請輸入E-mail')this.value='';"></td>
                          </tr>
<% 
	end if
%> 
		      </td>
                    </tr>
	<tr>
           <td colspan="5" class="IntroContent">請填寫下列問題：</td>
         </tr>
<%
	sql = "select * from m012 where m012_subjectid = " & subjectid & " order by m012_questionid"
	set rs2 = conn.execute(sql)
	i = 0
	while not rs2.eof
		i = i + 1
		css_str = ""
		if (i mod 2) = 1 then	css_str = " id=""tl_td"""
		
		AnswerType = trim(rs2("m012_Type"))
		if AnswerType = "1" or AnswerType = "" then
			sel = ""
			seltype = "radio"
		else
			sel = "（本題為複選題）"
			seltype = "checkbox"
		end if
%>
                    <tr bgcolor="dddddd">
                      <td align="center" valign="middle" class="IntroContent"><img src="images/Service/Q.gif" width="20" alt="問題圖示"></td>
                      <td colspan="4" class="IntroContent"><%= rs2("m012_questionid") %>. <%= trim(rs2("m012_title")) %><%= sel %></td>
                    </tr>
                    <tr bgcolor="eeeeee">
                      <td align="center" valign="middle" class="IntroContent"><img src="images/Service/A.gif" width="20" alt="答案圖示"></td>
                      <td colspan="4" valign="top" bgcolor="eeeeee" class="IntroContent">
<%
		sql = "" & _
			" select * from m013 where m013_subjectid = " & subjectid & _
			" and m013_questionid = " & rs2("m012_questionid") & _
			" order by m013_answerid "
		set rs3 = conn.execute(sql)
		while not rs3.eof
%>                            
        <label>
          <input type="<%= seltype %>" name="answer<%= rs2("m012_questionid") %>" value="<%= rs3("m013_answerid") %>"<% if trim(rs3("m013_default")) <> "" then %> checked<% end if %>>
	  <%= trim(rs3("m013_title")) %>
        </label>
<%
			if instr(open_str, "*" & rs2("m012_questionid") & "*" & rs3("m013_answerid") & "*") > 0 then
				response.write "<input type='text' name='open_content" & rs2("m012_questionid") & "_" & rs3("m013_answerid") & "' size='30' maxlength='512'>"
			end if
%>
                        <br>
<%
			rs3.movenext
		wend
%> 
	              </td>
                    </tr>
<%
		rs2.movenext
	wend
%>
                    <tr bgcolor="eeeeee">
                      <td class="IntroContent">其他意見</td>
                      <td><textarea name="reply" cols="50" rows="5" class="inputbox" colspan="4"></textarea></td>
                    </tr>
                    <tr>
                      <td height="50" align="center" bgcolor="eeeeee" colspan="5">
                        <input type=submit name=Submit value="完成">
                      </td>
                    </tr>
</table>
<hr size="1" />
</form>
</div>
