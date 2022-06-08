<%@ CodePage = 65001 %>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbFunc.inc" -->
<%
	ctNode = replace(trim(request("ctNode")), "'", "''")
	subjectid = replace(trim(request("subjectid")), "'", "''")
	if subjectid = "" then
		response.write "<script language='javascript'>alert('Error');history.go(-1);</script>"
	else

	set rs = conn.execute("select m011_onlyonce from m011 where m011_subjectid = " & subjectid)
	if rs.eof then
		response.write "<script language='javascript'>alert('Error');history.go(-1);</script>"
	else
	
	onlyonce = rs(0)

	m014_A = replace(request("m014_A"), "'", "''")
	m014_B = replace(request("m014_B"), "'", "''")
	m014_C = replace(request("m014_C"), "'", "''")
	m014_D = replace(request("m014_D"), "'", "''")
	m014_E = replace(request("m014_E"), "'", "''")
	m014_F = replace(request("m014_F"), "'", "''")
	m014_G = replace(request("m014_F"), "'", "''")
	m014_H = replace(request("m014_H"), "'", "''")

	reply = replace(request("reply"), "'", "''")


	if onlyonce = "1" and email <> "" then
		sql = "select mm014_A from m014 where m014_subjectid = " & subjectid & " and m014_A= '" & em014_A & "'"
		set rs = conn.execute(sql)
		if not rs.eof then
			response.write "<script language='javascript'>alert('你已做過此調查！');history.go(-1);</script>"
		end if
	else
	
	
	' 開放式答題
	sql = "select m015_questionid, m015_answerid from m015 where m015_subjectid = " & subjectid
	set rs3 = conn.execute(sql)
	while not rs3.eof
		open_str = open_str & "*" & rs3(0) & "*" & rs3(1) & "*,"
		open_str1 = open_str & "*" & rs3(0) & "* " & rs3(1) & "*,"
		rs3.movenext
'		response.write open_str
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


	sql2 = "select m012_questionid from m012 where m012_subjectid = " & subjectid & " order by 1"
	set rs = conn.execute(sql2)
	while not rs.eof
		questionid = rs(0)
		if request("answer" & questionid) = "" then
			response.write "<script language='javascript'>alert('請選擇第 " & questionid & " 題答案！');history.go(-1);</script>"
		else
			answer_a = split(request("answer" & questionid), ",")
			for i = 0 to ubound(answer_a)
				sql = sql & _
					" update m013 set m013_no = m013_no + 1 where " & _
					" m013_subjectid = " & subjectid & " and " & _
					" m013_questionid = " & questionid & " and " & _
					" m013_answerid = " & answer_a(i) & "; "
					
'					response.write answer_a(i) & ","
'					response.write questionid & ","
'					response.write trim(answer_a(i))
'					response.write open_str & ","
'					response.write "*" & questionid & "*" & trim(answer_a(i)) & "*" & ";"
'					response.write instr(open_str, "*" & questionid & "*" & trim(answer_a(i)) & "*") & ";"
'					response.write request("open_content" & questionid & "_" & trim(answer_a(i)))
'					response.write "</br>"

				if instr(open_str, "*" & questionid & "*" & trim(answer_a(i)) & "*") > 0 or instr(open_str1, "*" & questionid & "*" & answer_a(i) & "*") > 0 and trim(request("open_content" & questionid & "_" & answer_a(i))) <> "" then
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
						" '" & replace(trim(request("open_content" & questionid & "_" & trim(answer_a(i)))), "'", "''") & "', " & _
						" getdate() " & _
						" ); "
				else
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
						" '<null>', " & _
						" getdate() " & _
						" ); "		
				end if
			next
		end if
		rs.movenext
	wend
	conn.execute(sql)

        if request("mailid")<>"" then
           sql="update prosecute set checkflag=1 where id=" & request("mailid")
           conn.Execute(sql)
        end if 

	response.write "<script language='javascript'>alert('您的答題資料已經送出！');location.replace('sp.asp?xdURL=sVote/vote02.asp&subjectid=" & subjectid & "&ctNode=" & ctNode & "');</script>"
	end if
	end if
	end if
%>