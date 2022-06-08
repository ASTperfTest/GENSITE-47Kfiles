<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="意見信箱"
HTProgFunc="意見信箱維護"
HTProgCode="BM010"
HTProgPrefix="mSession"
response.expires = 0
%>
<!-- #include virtual = "/inc/server.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")

	id = request("id")


	if request("submit") = "刪除" then
		sql2 = "delete Prosecute where id = " & replace(Id, "'", "''")
		conn.execute(sql2)
		response.write "<script language='javascript'>alert('刪除完成!');location.replace('index.asp');</script>"
		response.end
	end if
	
	
	sql2 = "" & _
		" update Prosecute set " & _
		" reply = '" & replace(trim(request("reply")), "'", "''") & "', " & _
		" replydate = getdate(), " & _
		" sendflag = '' " & _
		" where id = " & replace(Id, "'", "''")
	conn.execute(sql2)
	if request("submit") = "儲存內容" then
		response.write "<script language='javascript'>alert('Down!');location.replace('index.asp');</script>"
		response.end
	end if


	set rs2 = conn.execute("select email, title from mailbox")
	if not rs2.eof then
		bossemail = trim(rs2("email"))
		title = trim(rs2("title"))
	end if

	Set mail = Server.CreateObject("CDONTS.NewMail")
	mail.subject = "" & session("mySiteName") & "意見信箱回覆"
	mail.from = bossemail
	mail.to = trim(request("email"))
	
	if trim(request("reply_cc")) <> "" then	mail.cc = trim(request("reply_cc"))
	mail.body = trim(request("reply"))	
	mail.send
	conn.execute("update prosecute set sendflag = '1' where id = " & replace(Id, "'", "''"))

	response.write "<script language='javascript'>alert('Down!');location.replace('index.asp');</script>"
%>
