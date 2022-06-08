<%@ CodePage = 65001 %>
<%
	HTProgCode = "GW1_vote01"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
	subjectid = request("subjectid")
	sql = "select m011_subject, m011_questno, m011_notetype, m011_bdate, m011_edate from m011 where m011_subjectid = " & replace(subjectid, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then	response.end
	notetype = trim(rs("m011_notetype"))
	

'	Response.ContentType = "application/vnd.ms-excel"
	response.ContentType = "application/csv"
	Response.Charset= "utf-8"
	response.AddHeader "Content-Disposition", "filename=export.csv;"
	
	response.write "<table><tr>"

	response.write "調查對象分析" 
	response.write "</tr><tr>"
	response.write "調查主題：" & trim(rs("m011_subject")) & ""
	response.write "</tr><tr>"
	response.write "起訖時間：" & datevalue(rs("m011_bdate")) & " ~ " & datevalue(rs("m011_edate")) & "" 
	response.write "</tr><tr color='blue'>"
	response.write "<td>調查時間</td>"
	if mid(notetype, 1, 1) = "1" then
	response.write "<td>姓名</td>"
	end if
	if mid(notetype, 3, 1) = "1" then
	response.write "<td>電話</td>"
	end if
	if mid(notetype, 11, 1) = "1" then
	response.write "<td>身分</td>"
	end if
	if mid(notetype, 2, 1) = "1" then
	response.write "<td>email</td>"
	end if
	if mid(notetype, 4, 1) = "1" then
	response.write "<td>年齡</td>"
	end if
	if mid(notetype, 5, 1) = "1" then
	response.write "<td>住所</td>"
	end if
	if mid(notetype, 9, 1) = "1" then
	response.write "<td>教育</td>"
	end if
	if mid(notetype, 10, 1) = "1" then
	response.write "<td>就診醫院</td>"
	end if
	if mid(notetype, 12, 1) = "1" then
	response.write "<td>就診地區</td>"
	end if
	if mid(notetype, 6, 1) = "1" then
	response.write "<td>家庭成員</td>"
	end if
	if mid(notetype, 7, 1) = "1" then
	response.write "<td>收入</td>"
	end if
	if mid(notetype, 8, 1) = "1" then
	response.write "<td>職業</td>"
	end if
	
	gsql= "select * from m012 where m012_subjectid= " & subjectid & " order by m012_questionid "
	set GSreg = conn.execute(gsql)
	while not GSreg.eof
		response.write "<td>" & GSreg("m012_questionid") & "." & GSreg("m012_title") & "</td>"
	GSreg.moveNext
	wend
	response.write "<td>其他意見</td>"
	response.write "</tr>"
	 
	asql = "" & _
		" select m014_id, m014_name, m014_sex, m014_idnumber, m014_email, m014_age, m014_addrarea, " & _
		" m014_familymember, m014_money, m014_reply, m014_job, " & _
		" m014_edu, m014_eid, m014_hospital, m014_hospitalarea  from m014" & _
		" where m014_subjectid = " & subjectid & other_sql & " order by m014_id " 
	set list = conn.execute(asql)
	while not list.EOF
		tsql = "select top 1 m016_updatetime from m016 where  m016_userid = "& list("m014_id") &""
		set Stime = conn.execute(tsql)
	
		response.write "<tr>"
				response.write "<td>" & Stime("m016_updatetime") & "</td>"
			
			if mid(notetype, 1, 1) = "1" then
				response.write "<td>" & list("m014_name") & "</td>"
			end if
			if mid(notetype, 3, 1) = "1" then
				response.write "<td>" & list("m014_sex") & "</td>"
			end if
			
			if mid(notetype, 11, 1) = "1" then
				response.write "<td>" & list("m014_eid") & "</td>"
			end if
			
			if mid(notetype, 2, 1) = "1" then
				response.write "<td><a href='mailto:" & trim(list("m014_email")) & "'>" & trim(list("m014_email")) & "</a></td>"
			end if
			
			if mid(notetype, 4, 1) = "1" then
				response.write "<td>" & list("m014_age") & "</td>"
			end if
			
			if mid(notetype, 5, 1) = "1" then
				response.write "<td>" & list("m014_addrarea") & "</td>"
			end if
			
			if mid(notetype, 9, 1) = "1" then
				response.write "<td>" & list("m014_edu") & "</td>"
			end if
			
			if mid(notetype, 10, 1) = "1" then
				response.write "<td>" & list("m014_hospital") & "</td>"
			end if

			if mid(notetype, 12, 1) = "1" then
				response.write "<td>" & list("m014_hospitalarea") & "</td>"
			end if
						
			if mid(notetype, 6, 1) = "1" then
				response.write "<td>" & list("m014_familymember") & "</td>"
			end if
			
			if mid(notetype, 7, 1) = "1" then
				response.write "<td>" & list("m014_money") & "</td>"
			end if
			
			if mid(notetype, 8, 1) = "1" then
				response.write "<td>" & list("m014_job") & "</td>"
			end if
			msql = "select max(m012_questionid) as maxquestion from m012 where m012_subjectid =  " & subjectid
			set MSreg = conn.execute(msql)
			for j=1 to MSreg("maxquestion")
				fsql = "select * from m016 where m016_userid = "& list("m014_id") & " and m016_questionid = " & j & " order by m016_questionid "
				
				set FSreg = conn.execute(fSql)	
				response.write "<td>" 
					while not FSreg.eof
					hsql = "select m013_title from m013 where m013_subjectid = " & subjectid &" and m013_questionid=  " & j & " and m013_answerid = " & FSreg("m016_answerid")
					set HSreg = conn.execute(hSql)	
						if FSreg("m016_content") <> "<null>"  Then
							response.write FSreg("m016_answerid") & "." & HSreg("m013_title") & "："
							response.write FSreg("m016_content") & "  "
							response.write "<br/>"
						else
							response.write FSreg("m016_answerid") & "." & HSreg("m013_title") & "  "
							response.write "<br/>"
						end if	
					FSreg.moveNext
				wend
				response.write "</td>" 
			next
			response.write "<td>" & list("m014_reply") & "</td>"
			response.write "</tr>"

		list.MoveNext
	wend
	response.write "</table>"
%>
