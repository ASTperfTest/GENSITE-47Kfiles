<% 
	Response.Expires = 0
	HTProgCap = "討論專區"
	HTProgFunc = "討論專區維護"
	HTProgCode = "BBS010"
	HTProgPrefix = "mSession" 
	response.expires = 0 
%>
<!-- #include virtual = "/inc/server.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	
	sql = "select isnull(max(id), 0) + 1 from talk05"
	set rs = conn.execute(sql)
	id = rs(0)
	
	name = trim(request("name"))
	master = trim(request("master"))
	seq = trim(request("seq"))
	
	if name = "" or not isnumeric(seq) then
		response.write "<script language='javascript'>alert('請輸入分類名稱、分類排序請輸入數字！');history.go(-1);</script>"
		response.end
	end if

	sql = "" & _
		" insert into talk05 ( " & _
		" id, name, master, updatetime, seq " & _
		" ) values( " & _
		id & ", " & _
		" '" & replace(name, "'", "''") & "', " & _
		" '" & replace(master, "'", "''") & "', " & _
		" getdate(), " & _
		replace(seq, "'", "''") & _
		" ) "
	conn.execute(sql)
	response.redirect "1_1.asp"
%>
 