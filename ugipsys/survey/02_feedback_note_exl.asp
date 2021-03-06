<%
	HTProgCode = "GW1_vote01"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
	response.ContentType = "application/csv"
	response.AddHeader "Content-Disposition", "filename=export.csv;"
	
	subjectid = request("subjectid")
	
	sql = "" & _
		" select m011_subject, m011_questno, m011_bdate, m011_edate " & _
		" from m011 where m011_subjectid = " & replace(subjectid, "'", "''")
	set rs = conn.execute(sql)
	if rs.eof then
		response.end
	end if
	
	response.write """意見回覆""" & vbcrlf
	response.write """調查主題：" & trim(rs("m011_subject")) & """" & vbcrlf
	response.write """起訖時間：" & datevalue(rs("m011_bdate")) & " ~ " & datevalue(rs("m011_edate")) & """" & vbcrlf
	response.write """姓名"", ""email"", ""意見內容""" & vbcrlf

	sql = "" & _
		" select m014_name, m014_email, m014_reply " & _
		" from m014 " & _
		" where m014_subjectid = " & replace(subjectid, "'", "''") & _
		" and isNull(m014_reply, '') not like '' " & _
		" order by m014_polldate desc "
	set rs = conn.execute(sql)
	while not rs.eof	
		response.write """" & trim(rs("m014_name")) & """, """ & trim(rs("m014_email")) & """, """ & trim(rs("m014_reply")) & """" & vbcrlf
		rs.movenext
	wend
%>
