﻿<%@ CodePage = 65001 %>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbFunc.inc" -->
<%
	subjectid = replace(trim(request("subjectid")), "'", "''")
	ctNode = request("ctNode")
	if subjectid = "" then
		response.write "<script language='javascript'>alert('操作錯誤！');history.go(-1);</script>"
		response.end
	end if

	set rs = conn.execute("select m011_onlyonce from m011 where m011_subjectid = " & subjectid)
	if rs.eof then
		response.write "<script language='javascript'>alert('操作錯誤！《2》');history.go(-1);</script>"
		response.end
	end if
	onlyonce = rs(0)
	
	m014_A = replace(request("m014_A"), "'", "''")
	m014_B = replace(request("m014_B"), "'", "''")
	m014_C = replace(request("m014_C"), "'", "''")
	m014_D = replace(request("m014_D"), "'", "''")
	m014_E = replace(request("m014_E"), "'", "''")
	m014_F = replace(request("m014_F"), "'", "''")
	m014_G = replace(request("m014_G"), "'", "''")
	m014_H = replace(request("m014_H"), "'", "''")
	reply = replace(request("reply"), "'", "''")
	ans_no_id = trim(request("ans_no_id"))

'	response.write ans_no_id
'	response.end

	sql = "select m015_questionid, m015_answerid from m015 where m015_subjectid = " & subjectid
	set rs3 = conn.execute(sql)
	while not rs3.EOF
		open_str = open_str & "*" & rs3(0) & "*" & rs3(1) & "*,"
		rs3.movenext
'		response.write open_str & "<hr/>"
	wend

	sql = ""
		set rs3 = conn.execute("select isNull(max(m014_id), 0) + 1 from m014")
		m014_id = rs3(0)
		sql = "" & _
			" insert into m014 ( " & _
			" m014_id, " & _
			" m014_A, " & _
			" m014_B, " & _
			" m014_C, " & _
			" m014_D, " & _
			" m014_E, " & _
			" m014_F, " & _
			" m014_G, " & _
			" m014_H, " & _
			" m014_pflag, " & _
			" m014_reply, " & _
			" m014_subjectid, " & _
			" m014_polldate " & _
			" ) values ( " & _
			m014_id & ", " & _
			" '" & m014_A & "', " & _
			" '" & m014_B & "', " & _
			" '" & m014_C & "', " & _
			" '" & m014_D & "', " & _
			" '" & m014_E & "', " & _
			" '" & m014_F & "', " & _
			" '" & m014_G & "', " & _
			" '" & m014_H & "', " & _
			" '0', " & _
			" '" & reply & "', " & _
			subjectid & ", " & _
			" getdate() " & _
			" ); "
'			response.write sql & "<hr/>"
			conn.execute(sql)        

	qno = split(ans_no_id, ",")
	for i = 0 to ubound(qno)
		qid_array = split(qno(i), "_")
		for j = 0 to ubound(qid_array)
			questionid = qid_array(0)
			answerid = qid_array(1)
		next  
'	response.write trim(request("open_content" & questionid & "_" & answerid)) & "<hr/>"
'	response.write instr(open_str, "*" & questionid & "*" & answerid & "*") & "<hr/>"
'	response.write replace(trim(request("open_content" & questionid & "_" & answerid)), "'", "''") & "<hr/>"
		if instr(open_str, "*" & questionid & "*" & answerid & "*") > 0 and trim(request("open_content" & questionid & "_" & answerid)) <> "" then
			sql16 = " insert into m016 ( " & _
				" m016_subjectid, " & _
				" m016_questionid, " & _
				" m016_answerid, " & _
				" m016_userid, " & _
				" m016_content, " & _
				" m016_updatetime " & _
				" ) values ( " & _
				subjectid & ", " & _
				questionid & ", " & _
				answerid & ", " & _
				m014_id & ", " & _
				" '" & replace(trim(request("open_content" & questionid & "_" & answerid)), "'", "''") & "', " & _
				" getdate() " & _
				" ); "
'				response.write sql16 & "<hr/>"
				conn.execute(sql16)        
		else
			sql16 = " insert into m016 ( " & _
				" m016_subjectid, " & _
				" m016_questionid, " & _
				" m016_answerid, " & _
				" m016_userid, " & _
				" m016_content, " & _
				" m016_updatetime " & _
				" ) values ( " & _
				subjectid & ", " & _
				questionid & ", " & _
				answerid & ", " & _
				m014_id & ", " & _
				" '<null>', " & _
				" getdate() " & _
				" ); "	
'				response.write sql16 & "<hr/>"
				conn.execute(sql16)        
		end if
		sql13 = "update m013 set m013_no = m013_no + 1 where m013_subjectid = " & subjectid & " and m013_questionid = " & questionid & " and m013_answerid = " & answerid
'		response.write sql13 & "<hr/>"
		conn.execute(sql13)        
	next
'				response.end	
  

	response.write "<script language='javascript'>alert('您的答題資料已經送出！');location.replace('sp.asp?xdURL=sVote/vote02.asp&subjectid=" & subjectid & "&ctNode=" & ctNode & "');</script>"
'	response.write "]]></pHTML></hpMain>"
'	response.end
%>