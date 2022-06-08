<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn02M02" %>
<!--#include virtual = "/inc/server.inc" -->
<% 
	 CatalogueFrame = "180,*"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset cols="<%=CatalogueFrame%>" framespacing="1" border="1" id="CatalogueMenu">
  <frame name="Catalogue" target="ForumToc" scrolling="auto" src="Folder_toc.asp?ItemID=<%=ItemID%>">
  <frame name="ForumToc" scrolling="auto" target="ShowArticle">
 <noframes>
  <body>

  <p>此網頁使用框架,但是您的瀏覽器並不支援.</p>

  </body>
  </noframes>
</frameset>

</html>
