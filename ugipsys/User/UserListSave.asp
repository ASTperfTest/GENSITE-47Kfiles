<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->

<%
if request("submittask")="批次存檔" then
    for each x in request.form
    	if left(x,5)="ckbox" and request(x)<>"" then     
	       xn=mid(x,6)
	       SQL="Update InfoUser Set ugrpName='" & request("tfx_ugrpName") & "',ugrpId='" & request("sfx_ugrpID") & _
	       		"' Where userId='" & request("pfx_UserID"&xn) & "'"
'response.write SQL & "<br>"	       		         
		conn.execute(SQL)
	   end if	  
    next       
end if
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>查詢結果清單</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<script language=VBS>
	alert "批次存檔完成"
	window.navigate "UserQuery.asp"
</script>
</body>
</html>