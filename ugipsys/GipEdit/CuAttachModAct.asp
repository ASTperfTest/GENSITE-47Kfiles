<!--#include virtual = "/inc/dbutil.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<title></title>
</head>
<body>
<%
	if session("UserID")="" then
	
		Response.Write "<SCRIPT language=JavaScript>alert('No Permissions!');window.history.back();</SCRIPT>"
		Response.End
		
	else
	
		Set Conn = Server.CreateObject("ADODB.Connection")
		
		Conn.Open session("ODBCDSN")
		
		Dim id 
		
		id=request("iCuItemID") 
		
	if Not IsNumeric(id) then 
	
		Response.Write "<SCRIPT language=JavaScript>alert('Error!');window.history.back();</SCRIPT>"
		Response.End

	
	else
	
		iCuItemID = request("iCuItemID")
		
	end if 
		sel = request("sel")
	
		xsql = "UPDATE CuDTx7 SET bgMusic = '" & sel & "' WHERE giCuItem = " & pkstr(iCuItemID,"")
		conn.execute xsql	
	
		Response.Write "<SCRIPT language=JavaScript>alert('修改成功');window.location.href='CuAttachList.asp?icuitem=" & iCuItemID & "&T="& now() &"';</SCRIPT>"
	
	
	end if
%>
</body>
</html>