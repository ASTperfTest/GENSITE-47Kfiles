<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="局長/民意 信箱"
HTProgFunc="局長信箱維護"
HTProgCode="BM010"
HTProgPrefix="mSession"
response.expires = 0
%>
<!-- #include virtual = "/inc/server.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")

	id = request("id")
	dateline = request("yy") & "/" & request("mm") & "/" & request("dd")

	if request("submit") = "刪除" then
		sql2 = "delete Prosecute where id = " & replace(Id, "'", "''")
		conn.execute(sql2)
		response.write "<script language='javascript'>alert('刪除完成!');location.replace('index.asp');</script>"
	else
		sql2 = "" & _
			" update Prosecute set " & _
			" docid = '" & replace(trim(request("docid")), "'", "''") & "', " & _
			" unit = '" & request("datafrom1") & "', " & _
			" classname = '', "
		if isdate(dateline) then
			sql2 = sql2 & " dateline = '" & dateline & "' "
		else
			sql2 = sql2 & " dateline = null "
		end if
		sql2 = sql2 & " where id = " & replace(Id, "'", "''")
		conn.execute(sql2)
		response.write "<script language='javascript'>alert('修改完成!');location.replace('index.asp');</script>"
	end if
%>