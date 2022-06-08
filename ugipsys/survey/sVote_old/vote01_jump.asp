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
<font><a>歡迎使用問卷調查</a></font>

<form method="post" action="ap.asp?xdURL=svote/inquire_qa_jump2.asp&ctNode=<%=ctNode%>" name="post" onSubmit="return send()">
<!--<form method="post" action="http://eppsys.hyweb.com.tw/site/demo/wsxd/svote/inquire_qa_jump2.asp" name="post">-->
	      <input type="hidden" name="subjectid" value="<%= subjectid %>">
	      <input type="hidden" name="mailid" value="<%= mailid %>">
<input type="hidden" name="questionid" value="<%= rs2("m012_questionid") %>">
<table width="55%" align="center" cellspacing="0" id="table_list" summary="列表頁">
  <tr>
    <th colspan="2" width="100%"><font><% =trim(rs("m011_subject"))%></font></th>
  </tr>
	<tr>
           <td colspan="2" class="IntroContent">請填寫相關資料，以便進行統計分析。</td>
         </tr>
<%
	notetype = trim(rs("m011_notetype"))
	
	if mid(notetype, 1, 1) = "1" then
%>
                          <tr>
                            <td class="IntroTitle"><strong>姓名：</strong></td>
                            <td><input type="text" name="name"></td>
                          </tr>
<%
	else
		response.write "<input type=""hidden"" name=""name"" value=""none"">"
	end if
	
	if mid(notetype, 2, 1) = "1" then
%>
                          <tr>
                            <td class="IntroTitle"><strong>電子郵件信箱：</strong></td>
                            <td><input type="text" name="email"></td>
                          </tr>
<% 
	else
		response.write "<input type=""hidden"" name=""email"" value=""@."">"
	end if 
	
	if mid(notetype, 3, 1) = "1" then
%>
                          <tr>
                            <td class="IntroTitle"><strong>電話：</strong></td>
                            <td><input type="text" name="sex"></td>
                          </tr>
                          <!--tr>
                            <td class="IntroTitle"><strong>性別：</strong></td>
                            <td>
                              <input type="radio" name="sex" value="M" checked>男
                              <input type="radio" name="sex" value="F">女
                            </td>
                          </tr-->
<% 
	end if 
	
	if mid(notetype, 4, 1) = "1" then
%>
                          <tr>
                            <td class="IntroTitle"><strong>年齡：</strong></td>
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
                            <td class="IntroTitle"><strong>居住縣市：</strong></td>
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
                            <td class="IntroTitle"><strong>家庭成員人數：</strong></td>
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
                            <td class="IntroTitle"><strong>收入：</strong></td>
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
                            <td class="IntroTitle"><strong>職業：</strong></td>
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
                            <td class="IntroTitle"><strong>教育程度：</strong></td>
                            <td>
			      <select name="edu">
				<option value=1 selected>小學及小學以下</option>
				<option value=2>國(初)中</option>
				<option value=3>高中(職)</option>
				<option value=4>大學(專)</option>
				<option value=5>研究所以上</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
	
	if mid(notetype, 10, 1) = "1" then
%>
                          <tr>
                            <td class="IntroTitle"><strong>就診醫院：</strong></td>
                            <td><input type="text" name="hospital"></td>
                          </tr>
<%
	end if
'		response.write "<input type=""hidden"" name=""hospital"" value="""">"
		
	if mid(notetype, 12, 1) = "1" then
%>
                          <tr>
                            <td class="IntroTitle"><strong>就診地區：</strong></td>
                            <td><input type="text" name="hospitalarea"></td>
                          </tr>
<%
	end if
'		response.write "<input type=""hidden"" name=""hospitalarea"" value="""">"

	if mid(notetype, 11, 1) = "1" then
%>
                          <tr>
                            <td class="IntroTitle"><strong>身分：</strong></td>
                            <td>
			      <select name="eid">
				<option value=1 selected>到院住院病患</option>
				<option value=2>到院門診病患</option>
				<option value=3>病患家屬</option>
				<option value=4>社區民眾</option>
				<option value=5>其他</option>
			      </select>
                            </td>
                          </tr>
<%
	end if
%> 
                        </table>
		      </td>
                    </tr>
</table>

<table width="55%" align="center" cellspacing="0" id="table_list" summary="列表頁">
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
                    <tr bgcolor="dddddd">
                      <td align="center" width="20" class="IntroContent"><img src="images/Service/Q.gif" width="25" height="25" alt="問題圖示"></td>
                      
                      <td colspan="4" class="IntroContent"><%= rs2("m012_questionid") %>. <%= trim(rs2("m012_title")) %><%= sel %></td>
                    </tr>
                    <tr bgcolor="eeeeee">
                      <td align="center" valign="top" class="IntroContent"><img src="images/Service/A.gif" width="25" height="25" alt="答案圖示"></td>
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
  <tr>
    <td align="center" colspan="5">
      <input type="submit" name="submit" value="下一步">
    </td>
  </tr>
</table>
<hr size="1" />
</form>
