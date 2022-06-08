<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GC1AP2" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<% 
	If ForumType = "D" Then
	 CatalogueFrame = "0,*"
   Else
	 CatalogueFrame = "220,*"
   End IF

   If ArticleType = "A" Then
	 ForumTocFrame = "35,40%,*"
   Else
	 ForumTocFrame = "0,*,0"
   End IF
   
   ItemID = request("CtRootID")
   


%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset cols="<%=CatalogueFrame%>" framespacing="1" border="1" id="CatalogueMenu">
  <frame name="Catalogue" target="ForumToc" scrolling="auto" noresize src="Folder_toc.asp?ItemID=<%=ItemID%>">
    <frame name="ForumToc" scrolling="auto" target="ShowArticle">
  <noframes>
  <body>

  <p>此網頁使用框架,但是您的瀏覽器並不支援.</p>

  </body>
  </noframes>
</frameset>

</html>
