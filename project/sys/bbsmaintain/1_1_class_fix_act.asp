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

	action = trim(request("Submit"))
	id = trim(request("id"))
	master = trim(request("master"))
	seq = trim(request("seq"))
	name = trim(request("name"))
	

	if name = "" or not isnumeric(seq) then
		response.write "<script language='javascript'>alert('�п�J�����W�١B�����Ƨǽп�J�Ʀr�I');history.go(-1);</script>"
		response.end
	end if


	if action = "�ק�" then
		sql = "" & _
			" update talk05 set " & _
			" name = '" & replace(name, "'", "''") & "', " & _
			" master = '" & replace(master, "'", "''") & "', " & _
			" seq = " & pkstr(seq, "") & ", " & _
			" updatetime = getdate() " & _
			" where id = " & pkstr(id, "")
	elseif action = "�R��" then
		sql = "" & _
			" delete talk05 where id = " & pkstr(id, "") & "; " & _
			" delete article where bid = " & pkstr(subject_id, "")
	end if
	
	conn.execute(sql)
	response.redirect "1_1.asp"
%>
 