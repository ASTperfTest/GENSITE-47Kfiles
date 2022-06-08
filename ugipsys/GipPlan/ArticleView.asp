<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "Pn03M04" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<% ForumID = Request.QueryString("ForumID")
   ForumEs = Request.QueryString("ForumEs") %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>文章檢視</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset rows="238,*">
  <frame name="ForumToc" scrolling="auto" target="ShowArticle" src="Article_toc.asp?ItemID=<%=ItemID%>&ForumID=<%=ForumID%>&ForumEs=<%=ForumEs%>">
  <frame name="ShowArticle" scrolling="auto">
  <noframes>
  <body>

  <p>此網頁使用框架,但是您的瀏覽器並不支援.</p>

  </body>
  </noframes>
</frameset>

</html>
