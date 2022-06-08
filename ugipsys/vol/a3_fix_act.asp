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
   id=replace(request("id"),"'","''")
   name=replace(request("name"),"'","''")
   passwd=replace(request("password"),"'","''")
   grade=replace(request("grade"),"'","''")
   unit=replace(request("unit"),"'","''")
   birthday=replace(request("birthday"),"'","''")
   
   sql="update volitient set name='" & name & "',passwd='" & passwd & "',grade='" & grade & "',unit='" & unit & "',birthday='" & birthday & "' where id='" & id & "'"
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
	doneStr="志工資料更新完成！"
	alert(doneStr)
        document.location.href = "index03.asp"
</script>
</body>
</html>
