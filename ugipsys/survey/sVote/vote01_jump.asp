<%@ CodePage = 65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<HEAD>
<TITLE>title</TITLE>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
</HEAD>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbFunc.inc" -->
<%
	subjectid = request("subjectid")
	ctNode = request("ctNode")
	if subjectid = "" then
		response.write "<script language='javascript'>history.go(-1);</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	
	
	sql = "select m011_notetype, m011_subject from m011 where m011_subjectid = " & subjectid
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤！');history.go(-1);</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	notetype = trim(rs("m011_notetype"))
	subject = trim(rs("m011_subject"))
	
	
	set ts = conn.execute("select count(*) from m012 where m012_subjectid = " & subjectid)
	if ts(0) = 0 then
		response.write "<script language='javascript'>alert('無調查問題！');history.go(-1);</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	
	
	sql = "" & _
		" select '*' + ltrim(str(m015_questionid)) + '*' + ltrim(str(m015_answerid)) + '*' " & _
		" from m015 where m015_subjectid = " & subjectid
	set rs3 = conn.execute(sql)
	while not rs3.EOF
		open_str = open_str & rs3(0) & ","
		rs3.movenext
	wend
	
	
	sql2 = "select top 1 * from m012 where m012_subjectid = " & subjectid & " order by m012_questionid"
	set rs2 = conn.execute(sql2)
%>

<script LANGUAGE="JavaScript">
<!--
function send1() {
	var form = document.post;

	if ( form.m014_A.value == "") {
		alert("您忘了填寫陳情的時間了！");
		form.m014_A.focus();
		return false;
	}

	if ( form.e_mail.B == "" ) {
		alert("您忘了填寫機關收發文號了！");
		form.e_B.focus();
		return false;
	}

	if ( form.e_mail.C == "" ) {
		alert("您忘了填寫陳請事由了！");
		form.C.focus();
		return false;
	}
	return true;
}

function send() {
	var form = document.post;

	if ( form.m014_name.value == "") {
		alert("您忘了填寫姓名了！");
		form.m014_name.focus();
		return false;
	}

	if ( form.e_mail.value == "" ) {
		alert("您忘了填寫email了！");
		form.e_mail.focus();
		return false;
	}

	return true;
}
//-->
</script>


<!-- 項目 -->
<div class="pageNav">
<ul><li><a href="mp.asp">首頁</a></li> > <li><a href="lp.asp?ctNode=73&CtUnit=46&BaseDSD=17">問卷調查</a></li> > <li>參與調查</li></ul>
</div>
<div class="print">
	<span class="previousPage">
		<a href="javascript:history.go(-1);">回上一頁</a>
	</span>
</div>
</br></br>
<h1>問卷調查</h1>
<div class="surveycp">
<% if subjectid = 1672 then %>
<form method="post" action="sp.asp?xdURL=svote/inquire_qa_jump2.asp&ctNode=<%=ctNode%>" name="post" onSubmit="return send1()">
<% else %>
<form method="post" action="sp.asp?xdURL=svote/inquire_qa_jump2.asp&ctNode=<%=ctNode%>" name="post" onSubmit="return send()">
<% end if %>
<!--<form method="post" action="http://eppsys.hyweb.com.tw/site/demo/wsxd/svote/inquire_qa_jump2.asp" name="post">-->
	      <input type="hidden" name="subjectid" value="<%= subjectid %>">
	      <input type="hidden" name="mailid" value="<%= mailid %>">
<input type="hidden" name="questionid" value="<%= rs2("m012_questionid") %>">
<table cellspacing="0" class="surveyTable" summary="意見調查表格">
  <tr>
    <th colspan="2" width="100%"><font><% =trim(rs("m011_subject"))%></font></th>
  </tr>
<%
	if subjectid = 1672 then
%>
	<tr>
           <td colspan="2" class="IntroContent">請填寫陳情案件相關資訊</td>
         </tr>
<%
	notetype = trim(rs("m011_notetype"))
	
	if mid(notetype, 9, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="time">請問您提出陳情的時間：</label><font color="#FF0000">(必填)</font></strong></td>
                          </tr>
                          <tr>
                            <td class="answer"><input type="text" name="m014_A" size="10" value="請輸入日期" tabindex="1" id="time"
onClick="javascript:window.dateField=document.post.m014_A; calendar=window.open('/inc/calendar2_west.html','cal','WIDTH=240,HEIGHT=260'); calendar.focus();" 
onKeyUp="javascript:window.dateField=document.post.m014_A; calendar=window.open('/inc/calendar2_west.html','cal','WIDTH=240,HEIGHT=260'); calendar.focus();" onFocus="if(this.value=='請輸入日期')this.value='';"></td>
                          </tr>
                          <noscript>
			  本網頁使用SCRIPT，如果您的瀏覽器不支援SCRIPT，請直接輸入日期 
			  </noscript>
<%
	end if
	
	if mid(notetype, 10, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="number">機關收發文號：</label><font color="#FF0000">(必填)</font><font color="#0000FF"> [※註:請務必填寫十一位數字之發文文號] </font></strong></td>
                          </tr>
                          <tr>
                            <td class="answer"><input type="text" name="m014_B" maxlength="11" size="20" value="請輸入機關收發文號" tabindex="2" id="number" onFocus="if(this.value=='請輸入機關收發文號')this.value='';"></td></td>
                          </tr>
<% 
	end if 
	
	if mid(notetype, 11, 1) = "1" then
%>
                           <tr>
                            <td class="question"><strong><label for="reason">陳請事由：</label><font color="#FF0000">(必填)</font></strong></td>
                          </tr>
                          <tr>
                            <td class="answer"><input type="text" name="m014_C" size="60" value="請輸入陳請事由" tabindex="3" id="reason" onFocus="if(this.value=='請輸入陳請事由')this.value='';"></td>
                           </tr>
<% 
	end if 
%> 
	<tr>
           <td colspan="2" class="IntroContent">請填寫基本資料</td>
         </tr>
<%
	if mid(notetype, 5, 1) = "1" then
%>
                            <td class="question"><strong><label for="sex">性別：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
                              <input type="radio" name="m014_E" value="男" tabindex="4" id="sex" checked>男
                              <input type="radio" name="m014_E" value="女" tabindex="5" id="sex">女
                            </td>
<% 
	end if
	
	if mid(notetype, 6, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="age">年齡：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
			      <select name="m014_F">
				<option value="20歲以下" tabindex="6" id="age">20歲以下</option>
				<option value="21-29歲" tabindex="7" id="age">21-29歲</option>
				<option value="30-39歲" tabindex="8" id="age">30-39歲</option>
				<option value="40-49歲" tabindex="9" id="age">40-49歲</option>
				<option value="50-59歲" tabindex="10" id="age">50-59歲</option>
				<option value="60歲以上" tabindex="11" id="age">60歲以上</option>
			      </select>
                            </td>
                          </tr>
<%
	end if 
	
	if mid(notetype, 7, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="edu">教育程度：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
			      <select name="m014_G">
				<option value="國小及以下" tabindex="12" id="edu">國小及以下</option>
				<option value="國中" tabindex="13" id="edu">國中</option>
				<option value="高中職" tabindex="14" id="edu">高中職</option>
				<option value="專科" tabindex="15" id="edu">專科</option>
				<option value="大學及以上" tabindex="16" id="edu">大學及以上</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
	
	if mid(notetype, 8, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="job">職業：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
			      <select name="m014_H">
				<option value="軍公教" tabindex="17" id="job">軍公教</option>
				<option value="企業主" tabindex="18" id="job">企業主</option>
				<option value="一般企業職員" tabindex="19" id="job">一般企業職員</option>
				<option value="勞工" tabindex="20" id="job">勞工</option>
				<option value="家管、退休人員" tabindex="21" id="job">家管、退休人員</option>
				<option value="自由業(醫師、律師、會計師等)" tabindex="22" id="job">自由業(醫師、律師、會計師等)</option>
				<option value="無業" tabindex="23" id="job">無業</option>
				<option value="其他" tabindex="24" id="job">其他</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
%> 
<% else %>
	<tr>
           <td colspan="2" class="IntroContent">請填寫相關資料，以便進行統計分析。</td>
         </tr>
<%
	notetype = trim(rs("m011_notetype"))
	
	if mid(notetype, 1, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="name">姓名：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer"><input type="text" name="m014_A" value="請輸入姓名" tabindex="1" id="name"></td>
                          </tr>
<%
	else
		response.write "<input type=""hidden"" name=""name"" value=""none"">"
	end if
	
	if mid(notetype, 2, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="connect">聯絡方式：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer"><input type="text" name="m014_B" value="請輸入聯絡方式" tabindex="2" id="connect"></td>
                          </tr>
<% 
	else
		response.write "<input type=""hidden"" name=""email"" value=""@."">"
	end if 
	
	if mid(notetype, 3, 1) = "1" then
%>
                           <tr>
                            <td class="question"><strong><label for="phone">電話：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer"><input type="text" name="m014_C" value="請輸入電話" tabindex="3" id="phone"></td>
                           </tr>
<% 
	end if 
	
	if mid(notetype, 4, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="email">E-mail：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer"><input type="text" name="m014_D" value="請輸入E-mail" tabindex="4" id="email"></td>
                          </tr>
<% 
	end if
	
	if mid(notetype, 5, 1) = "1" then
%>
                            <td class="question"><strong><label for="sex">性別：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
                              <input type="radio" name="m014_E" value="男" tabindex="4" id="sex" checked>男
                              <input type="radio" name="m014_E" value="女" tabindex="5" id="sex">女
                            </td>
<% 
	end if
	
	if mid(notetype, 6, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="age">年齡：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
			      <select name="m014_F">
				<option value="20歲以下" tabindex="6" id="age">20歲以下</option>
				<option value="21-29歲" tabindex="7" id="age">21-29歲</option>
				<option value="30-39歲" tabindex="8" id="age">30-39歲</option>
				<option value="40-49歲" tabindex="9" id="age">40-49歲</option>
				<option value="50-59歲" tabindex="10" id="age">50-59歲</option>
				<option value="60歲以上" tabindex="11" id="age">60歲以上</option>
			      </select>
                            </td>
                          </tr>
<%
	end if 
	
	if mid(notetype, 7, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="edu">教育程度：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
			      <select name="m014_G">
				<option value="國小及以下" tabindex="12" id="edu">國小及以下</option>
				<option value="國中" tabindex="13" id="edu">國中</option>
				<option value="高中職" tabindex="14" id="edu">高中職</option>
				<option value="專科" tabindex="15" id="edu">專科</option>
				<option value="大學及以上" tabindex="16" id="edu">大學及以上</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
	
	if mid(notetype, 8, 1) = "1" then
%>
                          <tr>
                            <td class="question"><strong><label for="job">職業：</label></strong></td>
                          </tr>
                          <tr>
                            <td class="answer">
			      <select name="m014_H">
				<option value="軍公教" tabindex="17" id="job">軍公教</option>
				<option value="企業主" tabindex="18" id="job">企業主</option>
				<option value="一般企業職員" tabindex="19" id="job">一般企業職員</option>
				<option value="勞工" tabindex="20" id="job">勞工</option>
				<option value="家管、退休人員" tabindex="21" id="job">家管、退休人員</option>
				<option value="自由業(醫師、律師、會計師等)" tabindex="22" id="job">自由業(醫師、律師、會計師等)</option>
				<option value="無業" tabindex="23" id="job">無業</option>
				<option value="其他" tabindex="24" id="job">其他</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
%> 
<% end if %>
                        </table>
		      </td>
                    </tr>
</table>

<table cellspacing="0" class="surveyTable" summary="列表頁">
	<tr>
           <td colspan="5" class="IntroContent">請填寫下列問題：</td>
         </tr>
<%
	while not rs2.eof
		AnswerType = trim(rs2("m012_Type"))
		if AnswerType = "1" or AnswerType = "" then
			sel = ""
			seltype = "radio"
		else
			sel = "（本題為複選題）"
			seltype = "checkbox"
		end if
%>
                    <tr>
                      <!--td align="center" width="20" class="IntroContent"><img src="images/Service/Q.gif" width="25" height="25" alt="問題圖示"></td-->
                      
                      <td class="question"><%= rs2("m012_questionid") %>. <%= trim(rs2("m012_title")) %><%= sel %></td>
                    </tr>
                    <tr bgcolor="eeeeee">
                      <!--td align="center" valign="top" class="IntroContent"><img src="images/Service/A.gif" width="25" height="25" alt="答案圖示"></td-->
                      <td class="answer">
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
				response.write "			<input type='text' name='open_content" & rs2("m012_questionid") & "_" & rs3("m013_answerid") & "' size='30' maxlength='512'>"
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
  <tr>
    <td align="center" colspan="5">
      <input type="submit" name="submit" value="下一步">
    </td>
  </tr>
</table>
<hr size="1" />
</form>
</div>
