<%@ CodePage = 65001 %>
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
	
	name = trim(request("name"))
	email = trim(request("email"))
	sex = trim(request("sex"))
	age = trim(request("age"))
	addrarea = trim(request("addrarea"))
	member = trim(request("member"))
	money = trim(request("money"))
	Job = trim(request("Job"))
	edu = request("edu")
	eid = trim(request("eid"))
	hospital = trim(request("hospital"))
	hospitalarea = request("hospitalarea")
	reply = trim(request("reply"))
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

	if name <> "" then
		set rs3 = conn.execute("select isNull(max(m014_id), 0) + 1 from m014")
		m014_id = rs3(0)
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
			" m014_pflag, " & _
			" m014_reply, " & _
			" m014_subjectid, " & _
			" m014_polldate, " & _
			" m014_eid, " & _
			" m014_hospital, " & _
			" m014_hospitalarea " & _
			" ) values ( " & _
			m014_id & ", " & _
			" '" & name & "', " & _
			" '" & sex & "', " & _
			" '" & email & "', " & _
			" '" & age & "', " & _
			" '" & addrarea & "', " & _
			" '" & member & "', " & _
			" '" & money & "', " & _
			" '" & job & "', " & _
			" '" & edu & "', " & _
			" '0', " & _
			" '" & reply & "', " & _
			subjectid & ", " & _
			" getdate(), " & _
			" '" & eid & "', " & _
			" '" & hospital & "', " & _
			" '" & hospitalarea & "' " & _			
			" ); "
'			response.write sql & "<hr/>"
			conn.execute(sql)        
	else
		m014_id = 0
		
		
	end if
'	response.write m014_id
'	response.end
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
  

	response.write "<script language='javascript'>alert('您的答題資料已經送出！');location.replace('ap.asp?xdURL=sVote/vote02.asp&subjectid=" & subjectid & "&ctNode=" & ctNode & "');</script>"
'	response.write "]]></pHTML></hpMain>"
'	response.end
%>