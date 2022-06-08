<%
	Response.Expires = 0
	HTProgCode = "GW1_vote01"
	response.expires = 0
%>
<!-- #include file = "../../../../inc/server.inc" -->
<!-- #include file = "../../../../inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%
	subjectid = request("subjectid")
	questionid = request("questionid")
	answerid = request("answer" & questionid)

	if subjectid = "" or questionid = "" or answerid = "" then
		response.write "<script language='javascript'>history.go(-1);</script>"
		response.end
	end if


	m014_name = trim(request("m014_name"))
	e_mail = trim(request("e_mail"))
	sex = trim(request("sex"))
	age = trim(request("age"))
	AddrArea = trim(request("AddrArea"))
	Member = trim(request("Member"))
	Money = trim(request("Money"))
	Job = trim(request("Job"))


	set rs = conn.execute("select m011_onlyonce, m011_subject from m011 where m011_subjectid = " & subjectid)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤！');history.go(-1);</script>"
		response.end
	end if
	subject = trim(rs("m011_subject"))
	m011_onlyonce = rs("m011_onlyonce")
	if e_mail <> "" and m011_onlyonce = "1" then
		set ts = conn.execute("select count(*) from m014 where m014_subjectid = " & subjectid & " and m014_email = '" & e_mail & "'")
		if ts(0) > 0 then
			response.write "<script language='javascript'>alert('你已做過此調查!');history.go(-1);</script>"
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
		response.end
	end if
	sql = "select m013_nextord from m013 where m013_subjectid = " & subjectid & " and m013_questionid = " & questionid & " and m013_answerid = " & answerid
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤(2)！');history.go(-1);</script>"
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
<html>
<head>
<title>問卷調查</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script LANGUAGE="JavaScript">
<!--
ie4 = ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4 ));
NN4 = (navigator.appName == "Netscape" && parseInt(navigator.appVersion) >= 4)
z = "xyz";
function toggle(targetId) {
	if ( NN4 || ie4 ) {
		if ( z != "xyz" ) {
			if ( z != document.all(targetId) )
				z.style.display = "none";
		}
		target = document.all(targetId);
		if ( target.style.display == "none" ) {
			target.style.display = "";
			z = document.all(targetId);
		} else {
			target.style.display = "none";
			z = "xyz";
		}
	}
}
ie4 = ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4 ));
NN4 = (navigator.appName == "Netscape" && parseInt(navigator.appVersion) >= 4)
z = "xyz";
//-->
</script>

<link href="template.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">

  <tr>
    <td>
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>

          <td align="center" valign="top"> 
            <table width="90%" border="0" cellpadding="0" cellspacing="1" style="margin-top:20px;">
              <tr>
                <td height="20"> 
                  <table width="100%" height="20" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="16" valign="middle" bgcolor="A4D265"><img src="../images/Service/service05.gif" width="16" height="20"></td>
                      <td width="80" valign="middle" bgcolor="A4D265" class="ContentTitle">問卷調查</td>
                      <td valign="middle" bgcolor="8DC73F" class="ContentTitle">&nbsp;</td>
                      <td width="16" align="right" bgcolor="8DC73F"><img src="../images/Service/service07.gif" width="16" height="20"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td height="40" bgcolor="eeeeee" class="SubTitle">歡迎使用問卷調查</td>
              </tr>
              <tr>
                <td height="80" align="center">
                  <table width="95%" border="0" align="center" cellpadding="3" cellspacing="2" class="px13">
<%
	if next_questionid = "0" then
%>
		    <form method="post" action="inquire_qa_act2.asp" name="post">
		    <input type="hidden" name="subjectid" value="<%= subjectid %>">
		    <input type="hidden" name="ans_no_id" value="<%= ans_no_id %>">
		    <input type="hidden" name="m014_name" value="<%= m014_name %>">
		    <input type="hidden" name="e_mail" value="<%= e_mail %>">
		    <input type="hidden" name="sex" value="<%= sex %>">
		    <input type="hidden" name="age" value="<%= age %>">
		    <input type="hidden" name="AddrArea" value="<%= AddrArea %>">
		    <input type="hidden" name="Member" value="<%= Member %>">
		    <input type="hidden" name="Money" value="<%= Money %>">
		    <input type="hidden" name="Job" value="<%= Job %>">
		    <%= open_content %>
                    <tr>
                      <td class="IntroTitle"><strong>其他意見：</strong></td>
                      <td><textarea name="reply" cols="50" rows="12" class="inputbox"></textarea></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td><input type="submit" name="submit" value="完成"></td>
                    </tr>
                    </form>
<%
	else
%>
		    <form method="post" action="inquire_qa_jump2.asp" name="post">
		    <input type="hidden" name="subjectid" value="<%= subjectid %>">
		    <input type="hidden" name="questionid" value="<%= next_questionid %>">
		    <input type="hidden" name="ans_no_id" value="<%= ans_no_id %>">
		    <input type="hidden" name="m014_name" value="<%= m014_name %>">
		    <input type="hidden" name="e_mail" value="<%= e_mail %>">
		    <input type="hidden" name="sex" value="<%= sex %>">
		    <input type="hidden" name="age" value="<%= age %>">
		    <input type="hidden" name="AddrArea" value="<%= AddrArea %>">
		    <input type="hidden" name="Member" value="<%= Member %>">
		    <input type="hidden" name="Money" value="<%= Money %>">
		    <input type="hidden" name="Job" value="<%= Job %>">
		    <%= open_content %>
                    <tr bgcolor="666666">
                      <td height="25" colspan="5" class="IntroTitle"><font color="ffffff"><%= subject %></font></td>
                    </tr>
<%
		sql = "select m012_type, m012_title from m012 where m012_subjectid = " & subjectid & " and m012_questionid = " & next_questionid
		set rs2 = conn.execute(sql)
		if rs2.eof then
			response.write "<script language='javascript'>alert('操作錯誤(3)！');history.go(-1);</script>"
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
                    <tr bgcolor="dddddd">
                      <td align="center" class="IntroContent"><img src="../images/Service/Q.gif" width="25" height="25"></td>
                      <td colspan="4" class="IntroContent"><%= next_questionid %>. <%= trim(rs2("m012_title")) %><%= sel %></td>
                    </tr>
                    <tr bgcolor="eeeeee">
                      <td align="center" valign="top" class="IntroContent"><img src="../images/Service/A.gif" width="25" height="25" alt="答案圖示"></td>
                      <td colspan="4" valign="top" bgcolor="eeeeee" class="IntroContent">
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
                    <tr bgcolor="#CDE1F1">
                      <td align="center" class="p12" colspan="5">
                        <input type="submit" name="submit" value="下一步">
                      </td>
                    </tr>
                    </form>
<%
	end if
%>
		  </table>
                </td>
              </tr>

            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
