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
<html>
<head>
<title>�ݨ��լd</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script LANGUAGE="JavaScript">
<!--
function send() {
	var form = document.post;
	if ( form.name.value == "" ) {
		alert("�п�J�m�W�I");
		form.name.focus();
		return false;
	} else if ( form.email.value == "" ) {
		alert("�п�Jemail�I");
		form.email.focus();
		return false;
	}

	return true;
}

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
<noscript>
  �z���s�������䴩script
</noscript>
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
                      <td width="16" valign="middle" bgcolor="A4D265"><img src="../images/Service/service05.gif" width="16" height="20" alt="�˹���"></td>
                      <td width="80" valign="middle" bgcolor="A4D265" class="ContentTitle">�ݨ��լd</td>
                      <td valign="middle" bgcolor="8DC73F" class="ContentTitle">&nbsp;</td>
                      <td width="16" align="right" bgcolor="8DC73F"><img src="../images/Service/service07.gif" width="16" height="20" alt="�˹���"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td height="40" bgcolor="eeeeee" class="SubTitle">�w��ϥΰݨ��լd</td>
              </tr>

<%
	subjectid = replace(trim(request("subjectid")), "'", "''")
	if subjectid = "" then
		response.write "<body onload=JavaScript:alert('�ާ@���~�I');history.go(-1);>"
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

	      <form method="post" action="inquire_qa_act.asp" name="post" onSubmit="javascript:return send();">
	      <input type="hidden" name="subjectid" value="<%= subjectid %>">
              <tr>
                <td height="80" align="center">
		  <table width="100%" border="0" cellspacing="1" cellpadding="2">
                    <tr bgcolor="666666">
                      <td height="25" colspan="5" class="IntroTitle"><font color="ffffff"><% =trim(rs("m011_subject"))%></font></td>
                    </tr>
                    <tr bgcolor="eeeeee">
                      <td colspan="5" align="center" class="IntroContent">
                        <table width="96%" border="0" cellspacing="4" cellpadding="0">
                          <tr>
                            <td height="20" colspan="2" class="IntroContent">�ж�g������ơA�H�K�i��έp���R�C</td>
                          </tr>
<%
	notetype = trim(rs("m011_notetype"))
	
	if mid(notetype, 1, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle">�m�W�G</td>
                            <td><input type="text" name="name"></td>
                          </tr>
<%
	else
		response.write "<input type=""hidden"" name=""name"" value=""none"">"
	end if
	
	if mid(notetype, 2, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle">�q�l�l��H�c�G</td>
                            <td><input type="text" name="email"></td>
                          </tr>
<% 
	else
		response.write "<input type=""hidden"" name=""email"" value=""@."">"
	end if 
	
	if mid(notetype, 3, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>�ʧO�G</strong></td>
                            <td>
                              <input type="radio" name="sex" value="M" checked>�k
                              <input type="radio" name="sex" value="F">�k
                            </td>
                          </tr>
<% 
	end if 
	
	if mid(notetype, 4, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>�~�֡G</strong></td>
                            <td>
			      <select name="age">
				<option value=1>0��11�� </option>
				<option value=2 selected>12��17��</option>
				<option value=3>18��23��</option>
				<option value=4>24��29��</option>
				<option value=5>30��39��</option>
				<option value=6>40��49��</option>
				<option value=7>50��59��</option>
				<option value=8>60���H�W</option>
			      </select>
                            </td>
                          </tr>
<% 
	end if
	
	if mid(notetype, 5, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>�~�����G</strong></td>
                            <td>
			      <select name="AddrArea">
				<option value=1 selected>�O�_��</option>
				<option value=2>�y����</option>
				<option value=3>��鿤</option>
				<option value=4>�s�˿�</option>
				<option value=5>�]�߿�</option>
				<option value=6>�O����</option>
				<option value=7>���ƿ�</option>
				<option value=8>�n�뿤</option>
				<option value=9>���L��</option>
				<option value=10>�Ÿq��</option>
				<option value=11>�O�n��</option>
				<option value=12>������</option>
				<option value=13>�̪F��</option>
				<option value=14>�O�F��</option>
				<option value=15>�Ὤ��</option>
				<option value=16>���</option>
				<option value=17>�򶩥�</option>
				<option value=18>�s�˥�</option>
				<option value=19>�O����</option>
				<option value=20>�Ÿq��</option>
				<option value=21>�O�n��</option>
				<option value=22>�O�_��</option>
				<option value=23>������</option>
				<option value=24>������</option>
				<option value=25>�s����</option>
			      </select>
                            </td>
                          </tr>
<% 
	end if
	
	if mid(notetype, 6, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>�a�x�����H�ơG</strong></td>
                            <td>
			      <select name="Member">
				<option value=1 selected>1�H</option>
				<option value=2>2�H</option>
				<option value=3>3�H</option>
				<option value=4>4�H</option>
				<option value=5>5��8�H</option>
				<option value=6>8�H�H�W</option>
			      </select>
                            </td>
                          </tr>
<%
	end if 
	
	if mid(notetype, 7, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>���J�G</strong></td>
                            <td>
			      <select name="Money">
				<option value=1 selected>5000���H�U</option>
				<option value=2>5000��1�U��</option>
				<option value=3>1�U��3�U��</option>
				<option value=4>3�U��5�U��</option>
				<option value=5>5�U��10�U��</option>
				<option value=6>�Q�U���H�W</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
	
	if mid(notetype, 8, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>¾�~�G</strong></td>
                            <td>
			      <select name="Job">
				<option value=1 selected>�x�H</option>
				<option value=2>���ȭ�</option>
				<option value=3>�ӤH</option>
				<option value=4>��¾</option>
				<option value=5>�ۥѷ~</option>
				<option value=6>�A�ȷ~</option>
				<option value=7>��L</option>
			      </select>
                            </td>
                          </tr>
<%
	end if

	if mid(notetype, 9, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>�Ǿ��G</strong></td>
                            <td>
			      <select name="edu">
				<option value=1 selected>��p</option>
				<option value=2>�ꤤ</option>
				<option value=3>����</option>
				<option value=4>�j��</option>
				<option value=5>�Ӥh</option>
				<option value=6>�դh</option>
				<option value=7>��L</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
%>
                        </table>
		      </td>
                    </tr>
<%
	sql = "select * from m012 where m012_subjectid = " & subjectid & " order by m012_questionid"
	set rs2 = conn.execute(sql)
	while not rs2.eof
		AnswerType = trim(rs2("m012_Type"))
		if AnswerType = "1" or AnswerType = "" then
			sel = ""
			seltype = "radio"
		else
			sel = "�]���D���ƿ��D�^"
			seltype = "checkbox"
		end if
%>
                    <tr bgcolor="dddddd">
                      <td align="center" class="IntroContent"><img src="../images/Service/Q.gif" width="25" height="25" alt="���D�ϥ�"></td>
                      <td colspan="4" class="IntroContent"><%= rs2("m012_questionid") %>. <%= trim(rs2("m012_title")) %><%= sel %></td>
                    </tr>
                    <tr bgcolor="eeeeee">
                      <td align="center" valign="top" class="IntroContent"><img src="../images/Service/A.gif" width="25" height="25" alt="���׹ϥ�"></td>
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
                    <tr bgcolor="eeeeee">
                      <td class="IntroContent">��L�N��</td>
                      <td><textarea name="reply" cols="50" rows="5" class="inputbox" colspan="4"></textarea></td>
                    </tr>
                    <tr>
                      <td height="50" align="center" bgcolor="eeeeee" colspan="5">
                        <input type=submit name=Submit value="����">
                      </td>
                    </tr>
                  </table>
                </td>
              </tr>
              </form>
              
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
