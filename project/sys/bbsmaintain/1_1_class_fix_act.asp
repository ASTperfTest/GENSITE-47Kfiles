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

	action = trim(request("Submit"))
	id = trim(request("id"))
	master = trim(request("master"))
	seq = trim(request("seq"))
	name = trim(request("name"))
	

	if name = "" or not isnumeric(seq) then
		response.write "<script language='javascript'>alert('請輸入分類名稱、分類排序請輸入數字！');history.go(-1);</script>"
		response.end
	end if


	if action = "修改" then
		sql = "" & _
			" update talk05 set " & _
			" name = '" & replace(name, "'", "''") & "', " & _
			" master = '" & replace(master, "'", "''") & "', " & _
			" seq = " & pkstr(seq, "") & ", " & _
			" updatetime = getdate() " & _
			" where id = " & pkstr(id, "")
	elseif action = "刪除" then
		sql = "" & _
			" delete talk05 where id = " & pkstr(id, "") & "; " & _
			" delete article where bid = " & pkstr(subject_id, "")
	end if
	
	conn.execute(sql)
	response.redirect "1_1.asp"
%>
 