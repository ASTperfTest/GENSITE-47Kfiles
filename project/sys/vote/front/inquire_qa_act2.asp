<%
	Response.Expires = 0
	'HTProgCap = "￥\¯aoT2z"
	'HTProgFunc = "°Y‥÷?O?d"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include file = "../../../../inc/server.inc" -->
<!-- #include file = "../../../../inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%
	subjectid = request("subjectid")
	ans_no_id = request("ans_no_id")
	
	if subjectid = "" or ans_no_id = "" then
		response.write "<script language='javascript'>history.go(-1);</script>"
		response.end
	end if
	

	username = trim(request("m014_name"))
	e_mail = trim(request("e_mail"))
	sex = trim(request("sex"))
	age = trim(request("age"))
	AddrArea = trim(request("AddrArea"))
	Member = trim(request("Member"))
	Money = trim(request("Money"))
	Job = trim(request("Job"))
	edu = request("edu")
	reply = trim(request("reply"))


	qno = split(ans_no_id, ",")
	for i = 0 to ubound(qno)
		qid_array = split(qno(i), "_")
		for j = 0 to ubound(qid_array)
			questionid = qid_array(0)
			answerid = qid_array(1)
		next  
		sql = "update m013 set m013_no = m013_no + 1 where m013_subjectid = " & subjectid & " and m013_questionid = " & questionid & " and m013_answerid = " & answerid
		conn.execute(sql)        
	next


	if username <> "" then
		set rs = conn.execute("select isNull(Max(m014_id), 0) + 1 from m014")
		sql = "" & _
			" insert into m014 ( " & _
			" m014_id, " & _
			" m014_name, " & _
			" m014_sex, " & _
			" m014_email, " & _
			" m014_age, " & _
			" m014_addrarea, " & _
			" m014_familymember, " & _
			" m014_money, " & _
			" m014_job, " & _
			" m014_edu, " & _
			" m014_reply, "  & _
			" m014_subjectid, " & _
			" m014_polldate " & _
			" ) values ( " & _
			rs(0) & ", " & _
			" '" & replace(username, "'", "''") & "', " & _
			" '" & replace(sex, "'", "''") & "', " & _
			" '" & replace(e_mail, "'", "''") & "', " & _
			" '" & replace(Age, "'", "''") & "', " & _
			" '" & replace(AddrArea, "'", "''") & "', " & _
			" '" & replace(Member, "'", "''") & "', " & _
			" '" & replace(Money, "'", "''") & "', " & _
			" '" & replace(Job, "'", "''") & "', " & _
			" '" & replace(edu, "'", "''") & "', " & _
			" '" & replace(reply, "'", "''") & "', " & _
			subjectid & ", " & _
			" getdate() " & _
			" ) "
		conn.execute(sql)
		
		
		sql = "select m015_questionid, m015_answerid from m015 where m015_subjectid = " & subjectid
		set rs3 = conn.execute(sql)
		while not rs3.EOF
			open_str = open_str & "*" & rs3(0) & "*" & rs3(1) & "*,"
			if trim(request("open_content" & rs3(0) & "_" & rs3(1))) <> "" then
				sql = "" & _
					" insert into m016 ( " & _
					" m016_subjectid, " & _
					" m016_questionid, " & _
					" m016_answerid, " & _
					" m016_userid, " & _
					" m016_content, " & _
					" m016_updatetime " & _
					" ) values ( " & _
					subjectid & ", " & _
					rs3(0) & ", " & _
					rs3(1) & ", " & _
					rs(0) & ", " & _
					" '" & replace(trim(request("open_content" & rs3(0) & "_" & rs3(1))), "'", "''") & "', " & _
					" getdate() " & _
					" ) "
				conn.execute(sql)
			end if
			rs3.movenext
		wend
	end if
  

	response.write "<script language='javascript'>alert('您的答題資料已經送出！');location.replace('vote02.asp?subjectid=" & subjectid & "');</script>"
	response.end
%>