<% 
	Response.Expires = 0
	HTProgCap = "�Q�ױM��"
	HTProgFunc = "�Q�ױM�Ϻ��@"
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
		response.write "<script language='javascript'>alert('�п�J�����W�١B�����Ƨǽп�J�Ʀr�I');history.go(-1);</script>"
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
 