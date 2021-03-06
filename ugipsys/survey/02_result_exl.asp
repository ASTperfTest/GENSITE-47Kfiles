<%@ CodePage = 65001 %>
<%
	HTProgCode = "GW1_vote01"
%>
<!--#include virtual = "/inc/server.inc" -->
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


	response.ContentType = "application/csv"
	response.AddHeader "Content-Disposition", "filename=export.csv;"

	
	response.write """答案累計統計""" & vbcrlf
	response.write """調查主題：" & trim(rs("m011_subject")) & """" & vbcrlf
	response.write """起訖時間：" & datevalue(rs("m011_bdate")) & " ~ " & datevalue(rs("m011_edate")) & """" & vbcrlf
	response.write """題目"", "
	for i = 1 to 30
		response.write """答案" & i & """, "
	next
	response.write vbcrlf
	
	
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
		
		response.write """第" & i & "題：" & trim(rs("m012_title")) & """, "
		
		sql = "" & _
			" select m013_title, isNull(m013_no, 0) " & _
			" from m013 where m013_subjectid = " & replace(subjectid, "'", "''") & _
			" and m013_questionid = " & questionid & _
			" order by m013_answerid "
		set rs2 = conn.execute(sql)
		while not rs2.eof
			response.write """" & trim(rs2("m013_title")) & """, "
			rs2.movenext
		wend
		response.write vbcrlf
		rs.movenext
	wend
	
	
	response.write """題號"", "
	for i = 1 to 30
		response.write """答案" & i & """, "
	next
	response.write vbcrlf

	
	
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
		
		response.write """第" & i & "題：" & trim(rs("m012_title")) & """, "
		
		sql = "" & _
			" select m013_title, isNull(m013_no, 0) no " & _
			" from m013 where m013_subjectid = " & replace(subjectid, "'", "''") & _
			" and m013_questionid = " & questionid & _
			" order by m013_answerid "
		set rs2 = conn.execute(sql)
		while not rs2.eof
			response.write """" & rs2("no") & """, "
			rs2.movenext
		wend
		response.write vbcrlf
		rs.movenext
	wend
%>
