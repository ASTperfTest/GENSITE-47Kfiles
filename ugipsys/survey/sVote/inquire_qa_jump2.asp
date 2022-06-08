<%@ CodePage = 65001 %>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbFunc.inc" -->
<%
	subjectid = request("subjectid")
	questionid = request("questionid")
	answerid = request("answer" & questionid)
	ctNode = request("ctNode")
'	response.write questionid & "<hr/>"
'	response.end

	if subjectid = "" then
		response.write "<script language='javascript'>history.go(-1);ppp</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	if questionid = "" then
		response.write "<script language='javascript'>history.go(-1);yyy</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	if answerid = "" then
		response.write "<script language='javascript'>history.go(-1);zzz</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if

	m014_A = replace(request("m014_A"), "'", "''")
	m014_B = replace(request("m014_B"), "'", "''")
	m014_C = replace(request("m014_C"), "'", "''")
	m014_D = replace(request("m014_D"), "'", "''")
	m014_E = replace(request("m014_E"), "'", "''")
	m014_F = replace(request("m014_F"), "'", "''")
	m014_G = replace(request("m014_G"), "'", "''")
	m014_H = replace(request("m014_H"), "'", "''")

	reply = replace(request("reply"), "'", "''")
	


	set rs = conn.execute("select m011_onlyonce, m011_subject from m011 where m011_subjectid = " & subjectid)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤！');history.go(-1);</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	subject = trim(rs("m011_subject"))
	m011_onlyonce = rs("m011_onlyonce")
	if e_mail <> "" and m011_onlyonce = "1" then
		set ts = conn.execute("select count(*) from m014 where m014_subjectid = " & subjectid & " and m014_email = '" & e_mail & "'")
		if ts(0) > 0 then
			response.write "<script language='javascript'>alert('你已做過此調查!');history.go(-1);</script>"
			response.write "]]></pHTML></hpMain>"
			response.end
		end if
	end if


	sql = "select m015_questionid, m015_answerid from m015 where m015_subjectid = " & subjectid
	set rs3 = conn.execute(sql)
	while not rs3.EOF
		open_str = open_str & "*" & rs3(0) & "*" & rs3(1) & "*,"
		if trim(request("open_content" & rs3(0) & "_" & rs3(1))) <> "" then
			open_content = open_content & "<input type='hidden' name='open_content" & rs3(0) & "_" & rs3(1) & "' value=""" & trim(request("open_content" & rs3(0) & "_" & rs3(1))) & """>"
		end if
		rs3.movenext
	wend


	if answerid = "" then
		response.write "<script language='javascript'>alert('請選擇答案！');history.go(-1);</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	sql = "select m013_nextord from m013 where m013_subjectid = " & subjectid & " and m013_questionid = " & questionid & " and m013_answerid = " & answerid
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤(2)！');history.go(-1);</script>"
		response.write "]]></pHTML></hpMain>"
		response.end
	end if
	next_questionid = rs(0)


	' 取得每次答題紀錄
	if request("ans_no_id") = "" then
		ans_no_id = questionid & "_" & answerid
	else
		ans_no_id = request("ans_no_id") & "," & questionid & "_" & answerid
	end if
%>


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
<table cellspacing="0" class="surveyTable" summary="意見調查表格">
  <tr class="text_12_000000">
    <td colspan="5"><div align="left"><%= subject %></div></td>
  </tr>
<%
	if next_questionid = "0" then
%>
  <form method="post" action="sp.asp?xdURL=svote/inquire_qa_act2.asp&ctNode=<%=ctNode%>" name="post">
  <!--<form method="post" action="http://socialsys.hyweb.com.tw/site/social/wsxdt/svote/inquire_qa_act2.asp" name="post">-->
	      <input type="hidden" name="subjectid" value="<%= subjectid %>">
	      <input type="hidden" name="mailid" value="<%= mailid %>">
  <input type="hidden" name="ans_no_id" value="<%= ans_no_id %>">
  <input type="hidden" name="m014_A" value="<%= m014_A %>">
  <input type="hidden" name="m014_B" value="<%= m014_B %>">
  <input type="hidden" name="m014_C" value="<%= m014_C %>">
  <input type="hidden" name="m014_D" value="<%= m014_D %>">
  <input type="hidden" name="m014_E" value="<%= m014_E %>">
  <input type="hidden" name="m014_F" value="<%= m014_F %>">
  <input type="hidden" name="m014_G" value="<%= m014_G %>">
  <input type="hidden" name="m014_H" value="<%= m014_H %>">
  <%= open_content %>
  <tr>
    <td class="question"><strong>其他意見：</strong></td>
  </tr>
  <tr>
    <td class="answer"><textarea name="reply" cols="50" rows="12" class="inputArea"></textarea></td>
  </tr>
  <tr>
    <td align="center" colspan="2"><input type="submit" name="submit" value="完成"></td>
  </tr>
  </form>
<%
	else
%>
  <form method="post" action="sp.asp?xdURL=svote/inquire_qa_jump2.asp&ctNode=<%=ctNode%>" name="post">
 	      <input type="hidden" name="subjectid" value="<%= subjectid %>">
	      <input type="hidden" name="mailid" value="<%= mailid %>">
  <input type="hidden" name="questionid" value="<%= next_questionid %>">
  <input type="hidden" name="ans_no_id" value="<%= ans_no_id %>">
  <input type="hidden" name="m014_A" value="<%= m014_A %>">
  <input type="hidden" name="m014_B" value="<%= m014_B %>">
  <input type="hidden" name="m014_C" value="<%= m014_C %>">
  <input type="hidden" name="m014_D" value="<%= m014_D %>">
  <input type="hidden" name="m014_E" value="<%= m014_E %>">
  <input type="hidden" name="m014_F" value="<%= m014_F %>">
  <input type="hidden" name="m014_G" value="<%= m014_G %>">
  <input type="hidden" name="m014_H" value="<%= m014_H %>">
  <%= open_content %>
<%
		sql = "select m012_type, m012_title from m012 where m012_subjectid = " & subjectid & " and m012_questionid = " & next_questionid
		set rs2 = conn.execute(sql)
		if rs2.eof then
			response.write "<script language='javascript'>alert('操作錯誤(3)！');history.go(-1);</script>"
			response.write "]]></pHTML></hpMain>"
			response.end
		end if

		answertype = trim(rs2("m012_type"))
		if answertype = "1" or answertype = "" then
			sel = ""
			seltype = "radio"
		else
			sel = "（本題為複選題）"
			seltype = "checkbox"
		end if
%>
                    <tr >
                      <td class="question"><%= next_questionid %>. <%= trim(rs2("m012_title")) %><%= sel %></td>
                    </tr>
                    <tr>
                      <td class="answer">
<%
		sql = "" & _
			" select * from m013 where m013_subjectid = " & subjectid & _
			" and m013_questionid = " & next_questionid & _
			" order by m013_answerid "
		set rs3 = conn.execute(sql)
		while not rs3.eof
%>                            
        <label>
          <input type="<%= seltype %>" name="answer<%= next_questionid %>" value="<%= rs3("m013_answerid") %>"<% if trim(rs3("m013_default")) <> "" then %> checked<% end if %>>
          <%= trim(rs3("m013_title")) %>
        </label>
<%
			if instr(open_str, "*" & next_questionid & "*" & rs3("m013_answerid") & "*") > 0 then
				response.write "			<input type='text' name='open_content" & next_questionid & "_" & rs3("m013_answerid") & "' size='30' maxlength='512'>"
			end if
%>
                        <br>
<%
			rs3.movenext
		wend
%> 
	              </td>
                    </tr> 
  <tr>
    <td align="center" colspan="5">
      <input type="submit" name="submit" value="下一步">
    </td>
  </tr>
</form>
<%
	end if
%>
</table>
</div>