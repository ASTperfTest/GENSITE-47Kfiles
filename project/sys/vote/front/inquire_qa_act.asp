<%
	Response.Expires = 0
	'HTProgCap = "�D\��aoT2z"
	'HTProgFunc = "�XY�L��?O?d"
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
	subjectid = replace(trim(request("subjectid")), "'", "''")
	if subjectid = "" then
		response.write "<script language='javascript'>alert('�ާ@���~�I');history.go(-1);</script>"
		response.end
	end if

	set rs = conn.execute("select m011_onlyonce from m011 where m011_subjectid = " & subjectid)
	if rs.eof then
		response.write "<script language='javascript'>alert('�ާ@���~�I�m2�n');history.go(-1);</script>"
		response.end
	end if
	onlyonce = rs(0)


	name = replace(request("name"), "'", "''")
	email= replace(request("email"), "'", "''")
	sex = request("sex")
	age = request("age")
	addrarea = request("addrarea")
	member = request("member")
	money = request("money")
	job = request("job")
	edu = request("edu")
	reply = replace(request("reply"), "'", "''")


	if onlyonce = "1" and email <> "" then
		sql = "select m014_email from m014 where m014_subjectid = " & subjectid & " and m014_email = '" & email & "'"
		set rs = conn.execute(sql)
		if not rs.eof then
			response.write "<script language='javascript'>alert('�A�w���L���լd�I');history.go(-1);</script>"
			response.end
		end if
	end if
	
	
	' �}�񦡵��D
	sql = "select m015_questionid, m015_answerid from m015 where m015_subjectid = " & subjectid
	set rs3 = conn.execute(sql)
	while not rs3.eof
		open_str = open_str & "*" & rs3(0) & "*" & rs3(1) & "*,"
		rs3.movenext
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
			" m014_polldate " & _
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
			" getdate() " & _
			" ); "
	else
		m014_id = 0
	end if


	sql2 = "select m012_questionid from m012 where m012_subjectid = " & subjectid & " order by 1"
	set rs = conn.execute(sql2)
	while not rs.eof
		questionid = rs(0)
		if request("answer" & questionid) = "" then
			response.write "<script language='javascript'>alert('�п�ܲ� " & questionid & " �D���סI');history.go(-1);</script>"
			response.end
		else
			answer_a = split(request("answer" & questionid), ",")
			for i = 0 to ubound(answer_a)
				sql = sql & _
					" update m013 set m013_no = m013_no + 1 where " & _
					" m013_subjectid = " & subjectid & " and " & _
					" m013_questionid = " & questionid & " and " & _
					" m013_answerid = " & answer_a(i) & "; "
					
				if instr(open_str, "*" & questionid & "*" & answer_a(i) & "*") > 0 and trim(request("open_content" & questionid & "_" & answer_a(i))) <> "" then
					sql = sql & _
						" insert into m016 ( " & _
						" m016_subjectid, " & _
						" m016_questionid, " & _
						" m016_answerid, " & _
						" m016_userid, " & _
						" m016_content, " & _
						" m016_updatetime " & _
						" ) values ( " & _
						subjectid & ", " & _
						questionid & ", " & _
						answer_a(i) & ", " & _
						m014_id & ", " & _
						" '" & replace(trim(request("open_content" & questionid & "_" & answer_a(i))), "'", "''") & "', " & _
						" getdate() " & _
						" ); "
				end if
			next
		end if
		rs.movenext
	wend
	'response.write sql
	'response.end
	conn.execute(sql)


	response.write "<script language='javascript'>alert('�z�����D��Ƥw�g�e�X�I');location.replace('vote02.asp?subjectid=" & subjectid & "');</script>"
	response.end
%>