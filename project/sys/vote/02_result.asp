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
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	
	sql = "" & _
		" select m011_subject, m011_bdate, m011_edate, m011_haveprize " & _
		" from m011 where m011_subjectid = " & replace(subjectid, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then
		response.write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><script language='JavaScript'>history.go(-1);</script></head><body></body></html>"
		response.end
	end if
	
	
	set ts = conn.execute("select count(*) from m014 where m014_subjectid = " & replace(subjectid, "'", "''"))
	if ts(0) = 0 then
		response.write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'><script language='JavaScript'>alert('目前尚無任何答題資料！');history.go(-1);</script></head><body></body></html>"
		response.end
	end if
	
	
	'開放式答題
	sql = "" & _
		" select '*' + ltrim(str(m015_questionid)) + '*' + ltrim(str(m015_answerid)) + '*' " & _
		" from m015 where m015_subjectid = " & replace(subjectid, "'", "''")
	set rs3 = conn.execute(sql)
	while not rs3.eof
		open_str = open_str & rs3(0) & ","
		rs3.movenext
	wend
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF">
<p>問卷調查</p>
<table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolorlight="#999900" bordercolordark="#FFFFFF">
  <tr> 
    <td>統計項目： 
      <input type="button" value="答案累計統計" onClick="javascript:location.replace('02_result.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'if rs("m011_haveprize") = "1" then
%>
      <input type="button" value="調查對象分析" onClick="javascript:location.replace('02_group.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
      <input type="button" value="調查對象明細" onClick="javascript:location.replace('02_namelist.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
<%
	'end if
%>
      <input type="button" value="意見回覆" onClick="javascript:location.replace('02_feedback_note.asp?functype=<%= functype %>&subjectid=<%= subjectid %>');">
    </td>
  </tr>

  <tr>
    <td>
      <input type="button" value="資料匯出" onClick="javascript:location.replace('02_result_exl.asp?subjectid=<%= subjectid %>');">
      <font color="#993300">[請按資料匯出之後,會出現另存新檔,副檔名請存成xls檔即可]</font>
    </td>
  </tr>
  <tr><td>起訖時間：<%= DateValue(rs("m011_bdate")) %> ~ <%= DateValue(rs("m011_edate")) %></td></tr>
  <tr><td>調查主題：<%= trim(rs("m011_subject")) %></tr>  

<%
	sql = "select m012_questionid, m012_title from m012 where m012_subjectid = " & subjectid & " order by 1"
	set rs2 = conn.execute(sql)
	i = 1
	while not rs2.EOF
%>
  <tr> 
    <td> 
      <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <tr><td colspan="3"><font color="blue">第 <%= i %> 題：<%= trim(rs2("m012_title")) %></font></td></tr>
<%
		sql = "select isNull(m013_no, 0) from m013 where m013_subjectid = " & subjectid & " and m013_questionid = " & rs2("m012_questionid")
		set rs3 = conn.execute(sql)             
		base_answer_no = 0
		while not rs3.EOF     
			base_answer_no = base_answer_no + rs3(0)
			rs3.movenext
		wend   
		if base_answer_no = 0 then
			response.write "<script language='JavaScript'>alert('目前尚無任何答題資料！');history.go(-1);</script>"
			response.end
		end if
                
		sql = "" & _
			" select m013_title, isNull(m013_no, 0) m013_no from m013 " & _
			" where m013_subjectid = " & subjectid & _
			" and m013_questionid = " & rs2("m012_questionid") & " order by m013_answerid "
		set rs3 = conn.execute(sql)
		j = 1
		while not rs3.EOF
			ratio = FormatPercent(rs3("m013_no") / base_answer_no, 1, False)
%>
        <tr> 
          <td>答案 <%= j %>：<%= trim(rs3("m013_title")) %><% if instr(open_str, "*" & i & "*" & j & "*") then %>　<a href="02_result_open.asp?subjectid=<%= subjectid %>&questionid=<%= i %>&answerid=<%= j %>">檢視開放式答題內容</a><% end if %></td>
          <td width="58%"><img src="images/bar_chart.gif" width="<%= ratio %>" height="12"></td>
          <td width="12%"><%= ratio %> (<%= rs3("m013_no") %>票)</td>
        </tr>
<%
			rs3.MoveNext
			j = j + 1
		wend
%>
      </table>      
    </td>
  </tr>
<%
		rs2.MoveNext
		i = i + 1
	wend
%>
  
</table>
<p><a href="javascript:history.go(-1);">Previous Page</a></p>
</body>
</html>
