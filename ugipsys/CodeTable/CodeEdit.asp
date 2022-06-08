<%@ CodePage = 65001 %>
<% 
response.expires=0
HTProgCap="代碼維護"
HTProgCode="Pn90M02"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches"> 
<meta http-equiv="Pragma" content="no-cache"> 
<title>代碼維護</title> 
<link rel="stylesheet" type="text/css" href="file:///D|/nmshygip/inc/setstyle.css"> 
</head>
<frameset id=f cols="35%,65%"> 
    <frame src="CodeStru.asp?myid=<%=request.querystring("myid")%>" name="contents" target="article"> 
    <frame src="CodeBlank.htm" name="article" target="contents"> 
</frameset><noframes></noframes> 
</html> 
