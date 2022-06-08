<%@ CodePage = 65001 %>
<% 
	Response.Expires = 0
   HTProgCode = "GC1AP5"
	 CatalogueFrame = "220,*"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%
if left(session("Public"),6)="/site/" and left(session("uploadPath"),6)<>"/site/" then
	session("uploadPath") = left(session("Public"),len(session("Public"))-8) & session("uploadPath")
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset cols="<%=CatalogueFrame%>" framespacing="1" border="1" id="CatalogueMenu">
  <frame name="Catalogue" target="ForumToc" scrolling="auto" noresize src="dirList.asp?xPath=<%=session("uploadPath")%>">
    <frame name="ForumToc" scrolling="auto" target="ShowArticle">
  <noframes>
  <body>

  <p>此網頁使用框架,但是您的瀏覽器並不支援.</p>

  </body>
  </noframes>
</frameset>

</html>
