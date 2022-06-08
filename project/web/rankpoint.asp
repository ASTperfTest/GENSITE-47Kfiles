<%
	Response.CacheControl = "no-cache" 
	Response.AddHeader "Pragma", "no-cache" 
	Response.Expires = -1
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<html xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:hyweb="urn:gip-hyweb-com" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>農業知識入口網 －小知識串成的大力量－/</title>
</head>

<!--#Include virtual = "/inc/client.inc" -->

<%

	sql = "select count(*) as total,memberid  from cudtgeneric inner join pedia on cudtgeneric.icuitem = pedia.gicuitem " & _
		"where ictunit = 2155 and fctupublic = 'Y' and xstatus = 'Y' and memberid in ( " & _
		"select distinct memberid from cudtgeneric inner join pedia on cudtgeneric.icuitem = pedia.gicuitem " & _
		"where fctupublic = 'Y' and xstatus = 'Y' and ictunit = 2155) group by memberid"
	set rs = conn.execute(sql)

response.write "<table border=""1"" width=""100%"">"
response.write "<caption>有推薦詞目且通過的會員ID,及其推薦數目</caption>"
response.write "<tr><td>"
	while not rs.eof
		
		if rs("memberId") <> "hyweb" then
		response.write "UPDATE ActivityPediaMember SET commendCount = " & rs("total") & " WHERE memberId = '" & rs("memberId") & "'; <br/>"
		end if
		rs.movenext
	wend
response.write "</td></tr><table>"
	rs.close
	set rs = nothing



	sql = "select count(*) as total,memberid  from cudtgeneric inner join pedia on cudtgeneric.icuitem = pedia.gicuitem " & _
		"where ictunit = 2156 and fctupublic = 'Y' and xstatus = 'Y' and memberid in ( " & _
		"select distinct memberid from cudtgeneric inner join pedia on cudtgeneric.icuitem = pedia.gicuitem " & _
		"where fctupublic = 'Y' and xstatus = 'Y' and ictunit = 2156) group by memberid"
	set rs = conn.execute(sql)

response.write "<table border=""1"" width=""100%"">"
response.write "<caption>有補充詞目且通過的會員ID,及其補充數目</caption>"
response.write "<tr><td>"
	while not rs.eof

		response.write "UPDATE ActivityPediaMember SET commendAdditionalCount = " & rs("total") & " WHERE memberId = '" & rs("memberId") & "'; <br/>"

		rs.movenext
	wend
response.write "</td></tr><table>"
	rs.close
	set rs = nothing
%>
