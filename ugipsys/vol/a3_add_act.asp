<% Response.Expires = 0
HTProgCap="志工資料"
HTProgFunc="志工資料"
HTProgCode="ap03"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<!--#include virtual = "/inc/hyftdGIP.inc" -->
<%
   id=replace(request("vid"),"'","''")
   name=replace(request("name"),"'","''")
   passwd=replace(request("password"),"'","''")
   grade=replace(request("grade"),"'","''")
   unit=replace(request("unit"),"'","''")
   birthday=replace(request("birthday"),"'","''")
   
   sql="select * from volitient where id='" & id & "'"
   set rs=conn.Execute(sql)
   if not rs.eof then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	doneStr="此志工編號已經存在！"
	alert(doneStr)
        document.location.href = "a3_add.asp"
</script>
</body>
</html>
<%     
   else
     sql="insert into volitient (id,name,passwd,grade,unit,birthday) values ('" & id & "','" & name & "','" & passwd & "','" & grade & "','" & unit & "','" & birthday & "')"
     conn.Execute(sql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	doneStr="志工資料新增完成！"
	alert(doneStr)
        document.location.href = "index03.asp"
</script>
</body>
</html>
<%     
   end if
%>