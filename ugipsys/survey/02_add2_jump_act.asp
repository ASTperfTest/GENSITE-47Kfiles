<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%	
	submit_str = trim(request("submit"))
	subjectid = request("subjectid")
	questno = request("questno")
	jumpquestion = request("jumpquestion")
	userid = request.cookies("id")
	open_str = "開放式答題"
	
	if subjectid = "" or questno = "" or jumpquestion = "" or submit_str = "" then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	
	
	
	if submit_str <> "重新調查此份問卷" and submit_str <> "確定" and submit_str <> "修改資料" then
		questionid = mid(submit_str, 5)
		sql = "" & _
			" update m012 set " & _
			" m012_title = '" & replace(trim(request("title" & questionid)), "'", "''") & "', " & _
			" m012_answerno = " & request("answerno" & questionid) & ", " & _
			" m012_type = '" & request("type" & questionid) & "', " & _
			" m012_modifyuser = '" & userid & "', " & _
			" m012_updatetime = getdate() " & _
			" where m012_subjectid = " & subjectid & " and m012_questionid = " & questionid & "; "
		
		for i = 1 to request("answerno" & questionid)
			sql = sql & _
				" update m013 set " & _
				" m013_title = '" & replace(trim(request("answer" & questionid & "-" & i)), "'", "''") & "', " & _
				" m013_default = '" & request("default" & questionid & "-" & i) & "', " & _
				" m013_nextord = " & request("nextord" & questionid & "-" & i) & ", " & _
				" m013_modifyuser = '" & userid & "', " & _
				" m013_updatetime = getdate() " & _
				" where m013_subjectid = " & subjectid & " and " & _
				" m013_questionid = " & questionid & " and " & _
				" m013_answerid = " & i & "; "
				
			if request("open" & questionid & "-" & i) = "Y" then
				sql = sql & _
					" insert into m015 ( " & _
					" m015_subjectid, " & _
					" m015_questionid, " & _
					" m015_answerid, " & _
					" m015_modifyuser, " & _
					" m015_updatetime " & _
					" ) values ( " & _
					subjectid & ", " & _
					questionid & ", " & _
					i & ", " & _
					" '" & userid & "', " & _
					" getdate() " & _
					" ); "
			end if
		next
		
		sql = "" & _
			" delete from m015 where m015_subjectid = " & subjectid & " and m015_questionid = " & questionid & "; " & _
			" delete from m016 where m016_subjectid = " & subjectid & " and m016_questionid = " & questionid & "; " & _
			sql
		conn.execute(sql)
		response.redirect "02.asp"
		response.end
	end if




	dim	answer_no(30, 30)
	for i = 0 to 30
		for j = 0 to 30
			answer_no(i, j) = 0
		next
	next
	
	if submit_str = "重新調查此份問卷"  or submit_str = "確定" then
		sql = "" & _
			" update m011 set m011_pflag = '0' where m011_subjectid = " & subjectid & ";" & _
			" delete from m012 where m012_subjectid = " & subjectid & "; " & _
			" delete from m013 where m013_subjectid = " & subjectid & "; " & _
			" delete from m014 where m014_subjectid = " & subjectid & "; " & _
			" delete from m015 where m015_subjectid = " & subjectid & "; " & _
			" delete from m016 where m016_subjectid = " & subjectid & "; "
	elseif submit_str = "修改資料" then
		sql = "" & _
			" select m013_questionid, m013_answerid, m013_no from m013 " & _
			" where m013_subjectid = " & subjectid & " order by m013_questionid, m013_answerid "
		set rs_no = conn.execute(sql)
		while not rs_no.EOF
			answer_no(rs_no("m013_questionid"), rs_no("m013_answerid")) = rs_no("m013_no")
			rs_no.MoveNext
		wend
		
		sql = "" & _
			" delete from m012 where m012_subjectid = " & subjectid & "; " & _
			" delete from m013 where m013_subjectid = " & subjectid & "; " & _
			" delete from m015 where m015_subjectid = " & subjectid & "; "
	end if
	conn.execute(sql)
	
	
	
	for i = 1 to questno
		
		sql_question_tmp = "" & _
			" insert into m012 ( " & _
			" m012_subjectid, " & _
			" m012_questionid, " & _
			" m012_title, " & _
			" m012_answerno, " & _
			" m012_type, " & _
			" m012_createuser, " & _
			" m012_createtime, " & _
			" m012_modifyuser, " & _
			" m012_updatetime " & _
			" ) values ( " & _
			subjectid & ", " & _
			i & ", " & _
			" '" & replace(trim(request("title" & i)), "'", "''") & "', " & _
			request("answerno" & i) & ", " & _
			" '" & request("type" & i) & "', " & _
			" '" & userid & "', " & _
			" getdate(), " & _
			" '" & userid & "', "& _
			" getdate() " & _
			" ); "
		sql_question = sql_question & sql_question_tmp
		
		
		for j = 1 to request("answerno" & i)
			sql_answer_tmp = "" & _
				" insert into m013 ( " & _
				" m013_subjectid, " & _
				" m013_questionid, " & _
				" m013_answerid, " & _
				" m013_title, " & _
				" m013_default, " & _
				" m013_nextord, " & _
				" m013_no, " & _
				" m013_createuser, " & _
				" m013_createtime, " & _
				" m013_modifyuser, " & _
				" m013_updatetime " & _
				" ) values ( " & _
				subjectid & ", " & _
				i & ", " & _
				j & ", " & _
				" '" & replace(trim(request("answer" & i & "-" & j)), "'", "''") & "', " & _
				" '" & request("default" & i & "-" & j) & "', " & _
				request("nextord" & i & "-" & j) & ", " & _
				answer_no(i, j) & ", " & _
				" '" & userid & "', " & _
				" getdate(), " & _
				" '" & userid & "', "& _
				" getdate() " & _
				" ); "
			sql_answer = sql_answer & sql_answer_tmp
			
			sql_open_tmp = ""
			if request("open" & i & "-" & j) = "Y" then
				sql_open_tmp = "" & _
					" insert into m015 ( " & _
					" m015_subjectid, " & _
					" m015_questionid, " & _
					" m015_answerid, " & _
					" m015_modifyuser, " & _
					" m015_updatetime " & _
					" ) values ( " & _
					subjectid & ", " & _
					i & ", " & _
					j & ", " & _
					" '" & userid & "', " & _
					" getdate() " & _
					" ); "
				sql_open = sql_open & sql_open_tmp
			end if
			
		next
	next
	
	'response.write sql_question & sql_answer & sql_open
	'response.end
	
	conn.execute(sql_question & sql_answer & sql_open)
	response.redirect "/GipEdit/DsdXMLList.asp"
%>