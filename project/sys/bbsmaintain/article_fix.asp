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
	
	
	submit_str = trim(request("submit"))
	bid = trim(request("bid"))
	mid2 = trim(request("mid"))
	sid = trim(request("sid"))
	if bid = "" or mid2 = "" or sid = "" then
		response.redirect "index.asp"
		response.end
	end if
  
    
	if submit_str = "修改" then
		sql = "" & _
			" update article set " & _
			" title = " & pkstr(request("title"), "") & ", " & _
			" message = " & pkstr(request("message"), "") & _
			" where sid = " & pkstr(sid, "") & _
			" and mid = " & pkstr(mid2, "") & " and bid = " & pkstr(bid, "")
	elseif submit_str = "刪除" then
		sql = "" & _
			" delete from article where sid = " & pkstr(sid, "") & _
			" and mid = " & pkstr(mid2, "") & " and bid = " & pkstr(bid, "")
	end if
	conn.execute(sql)
     
	response.redirect "articlelist.asp?bid=" & bid
%>