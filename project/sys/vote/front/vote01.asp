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
<title>問卷調查</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
  您的瀏覽器不支援script
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
                      <td width="16" valign="middle" bgcolor="A4D265"><img src="../images/Service/service05.gif" width="16" height="20" alt="裝飾圖"></td>
                      <td width="80" valign="middle" bgcolor="A4D265" class="ContentTitle">問卷調查</td>
                      <td valign="middle" bgcolor="8DC73F" class="ContentTitle">&nbsp;</td>
                      <td width="16" align="right" bgcolor="8DC73F"><img src="../images/Service/service07.gif" width="16" height="20" alt="裝飾圖"></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td height="40" bgcolor="eeeeee" class="SubTitle">歡迎使用問卷調查</td>
              </tr>

<%
	subjectid = replace(trim(request("subjectid")), "'", "''")
	if subjectid = "" then
		response.write "<body onload=JavaScript:alert('操作錯誤！');history.go(-1);>"
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
                            <td height="20" colspan="2" class="IntroContent">請填寫相關資料，以便進行統計分析。</td>
                          </tr>
<%
	notetype = trim(rs("m011_notetype"))
	
	if mid(notetype, 1, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle">姓名：</td>
                            <td><input type="text" name="name"></td>
                          </tr>
<%
	else
		response.write "<input type=""hidden"" name=""name"" value=""none"">"
	end if
	
	if mid(notetype, 2, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle">電子郵件信箱：</td>
                            <td><input type="text" name="email"></td>
                          </tr>
<% 
	else
		response.write "<input type=""hidden"" name=""email"" value=""@."">"
	end if 
	
	if mid(notetype, 3, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>性別：</strong></td>
                            <td>
                              <input type="radio" name="sex" value="M" checked>男
                              <input type="radio" name="sex" value="F">女
                            </td>
                          </tr>
<% 
	end if 
	
	if mid(notetype, 4, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>年齡：</strong></td>
                            <td>
			      <select name="age">
				<option value=1>0到11歲 </option>
				<option value=2 selected>12到17歲</option>
				<option value=3>18到23歲</option>
				<option value=4>24到29歲</option>
				<option value=5>30到39歲</option>
				<option value=6>40到49歲</option>
				<option value=7>50到59歲</option>
				<option value=8>60歲以上</option>
			      </select>
                            </td>
                          </tr>
<% 
	end if
	
	if mid(notetype, 5, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>居住縣市：</strong></td>
                            <td>
			      <select name="AddrArea">
				<option value=1 selected>臺北縣</option>
				<option value=2>宜蘭縣</option>
				<option value=3>桃園縣</option>
				<option value=4>新竹縣</option>
				<option value=5>苗栗縣</option>
				<option value=6>臺中縣</option>
				<option value=7>彰化縣</option>
				<option value=8>南投縣</option>
				<option value=9>雲林縣</option>
				<option value=10>嘉義縣</option>
				<option value=11>臺南縣</option>
				<option value=12>高雄縣</option>
				<option value=13>屏東縣</option>
				<option value=14>臺東縣</option>
				<option value=15>花蓮縣</option>
				<option value=16>澎湖縣</option>
				<option value=17>基隆市</option>
				<option value=18>新竹市</option>
				<option value=19>臺中市</option>
				<option value=20>嘉義市</option>
				<option value=21>臺南市</option>
				<option value=22>臺北市</option>
				<option value=23>高雄市</option>
				<option value=24>金門縣</option>
				<option value=25>連江縣</option>
			      </select>
                            </td>
                          </tr>
<% 
	end if
	
	if mid(notetype, 6, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>家庭成員人數：</strong></td>
                            <td>
			      <select name="Member">
				<option value=1 selected>1人</option>
				<option value=2>2人</option>
				<option value=3>3人</option>
				<option value=4>4人</option>
				<option value=5>5至8人</option>
				<option value=6>8人以上</option>
			      </select>
                            </td>
                          </tr>
<%
	end if 
	
	if mid(notetype, 7, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>收入：</strong></td>
                            <td>
			      <select name="Money">
				<option value=1 selected>5000元以下</option>
				<option value=2>5000到1萬元</option>
				<option value=3>1萬到3萬元</option>
				<option value=4>3萬到5萬元</option>
				<option value=5>5萬到10萬元</option>
				<option value=6>十萬元以上</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
	
	if mid(notetype, 8, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>職業：</strong></td>
                            <td>
			      <select name="Job">
				<option value=1 selected>軍人</option>
				<option value=2>公務員</option>
				<option value=3>商人</option>
				<option value=4>教職</option>
				<option value=5>自由業</option>
				<option value=6>服務業</option>
				<option value=7>其他</option>
			      </select>
                            </td>
                          </tr>
<%
	end if

	if mid(notetype, 9, 1) = "1" then
%>
                          <tr>
                            <td width="110" class="IntroTitle"><strong>學歷：</strong></td>
                            <td>
			      <select name="edu">
				<option value=1 selected>國小</option>
				<option value=2>國中</option>
				<option value=3>高中</option>
				<option value=4>大學</option>
				<option value=5>碩士</option>
				<option value=6>博士</option>
				<option value=7>其他</option>
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
			sel = "（本題為複選題）"
			seltype = "checkbox"
		end if
%>
                    <tr bgcolor="dddddd">
                      <td align="center" class="IntroContent"><img src="../images/Service/Q.gif" width="25" height="25" alt="問題圖示"></td>
                      <td colspan="4" class="IntroContent"><%= rs2("m012_questionid") %>. <%= trim(rs2("m012_title")) %><%= sel %></td>
                    </tr>
                    <tr bgcolor="eeeeee">
                      <td align="center" valign="top" class="IntroContent"><img src="../images/Service/A.gif" width="25" height="25" alt="答案圖示"></td>
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
                      <td class="IntroContent">其他意見</td>
                      <td><textarea name="reply" cols="50" rows="5" class="inputbox" colspan="4"></textarea></td>
                    </tr>
                    <tr>
                      <td height="50" align="center" bgcolor="eeeeee" colspan="5">
                        <input type=submit name=Submit value="完成">
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
