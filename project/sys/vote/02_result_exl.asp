<%@ CodePage = 65001 %>
<%
	HTProgCode = "GW1_vote01"
%>
<!-- #include virtual = "/inc/client.inc" -->
<%
	subjectid = request("subjectid")
	
	sql = "" & _
		" select m011_subject, m011_questno, m011_bdate, m011_edate " & _
		" from m011 where m011_subjectid = " & replace(subjectid, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then
		response.end
	end if


	questno = rs("m011_questno")


	response.ContentType = "application/vnd.ms-excel; charset=Big5"
	response.AddHeader "Content-Disposition", "filename=export.xls;"
	response.AddHeader "Accept-Language", "zh-tw"
	
	response.write "<table>"
	response.write "<tr><td colspan=81>答案累計統計</td></tr>" & vbcrlf
	response.write "<tr><td colspan=81>調查主題：" & trim(rs("m011_subject")) & "</td></tr>" & vbcrlf
	response.write "<tr><td colspan=81>起訖時間：" & datevalue(rs("m011_bdate")) & " ~ " & datevalue(rs("m011_edate")) & "</td></tr>" & vbcrlf
	response.write "<tr><td>題目</td> "
	for i = 1 to 80
		response.write "<td>答案" & i & "</td> "
	next
	response.write vbcrlf
	response.write "</tr>"
	
	sql = "" & _
		" select m012_questionid, m012_title, m012_answerno " & _
		" from m012 where m012_subjectid = " & replace(subjectid, "'", "''") & _
		" order by m012_questionid "
	set rs = conn.execute(sql)
	i = 0
	while not rs.eof
		i = i + 1
		questionid = rs("m012_questionid")
		answerno = rs("m012_answerno")
		
		response.write "<tr><td>第" & i & "題：" & trim(rs("m012_title")) & "</td> "
		
		sql = "" & _
			" select m013_title, isNull(m013_no, 0) " & _
			" from m013 where m013_subjectid = " & replace(subjectid, "'", "''") & _
			" and m013_questionid = " & questionid & _
			" order by m013_answerid "
		set rs2 = conn.execute(sql)
		while not rs2.eof
			response.write "<td>" & trim(rs2("m013_title")) & "</td> "
			rs2.movenext
		wend
		response.write vbcrlf
		response.write "</tr>"
	
		rs.movenext
	wend
	
	
	response.write "<tr><td>題號</td> "
	for i = 1 to 80
		response.write "<td>答案" & i & "</td> "
	next
	response.write vbcrlf
	response.write "</tr>"
	
	
	sql = "" & _
		" select m012_questionid, m012_title, m012_answerno " & _
		" from m012 where m012_subjectid = " & replace(subjectid, "'", "''") & _
		" order by m012_questionid "
	set rs = conn.execute(sql)
	i = 0
	while not rs.eof
		i = i + 1
		questionid = rs("m012_questionid")
		answerno = rs("m012_answerno")
		
		response.write "<tr><td>第" & i & "題：" & trim(rs("m012_title")) & "</td> "
		
		sql = "" & _
			" select m013_title, isNull(m013_no, 0) no " & _
			" from m013 where m013_subjectid = " & replace(subjectid, "'", "''") & _
			" and m013_questionid = " & questionid & _
			" order by m013_answerid "
		set rs2 = conn.execute(sql)
		while not rs2.eof
			response.write "<td>" & rs2("no") & "</td> "
			rs2.movenext
		wend
		response.write vbcrlf
		response.write "</tr>"
		rs.movenext
	wend
	response.write "</table>"
%>
