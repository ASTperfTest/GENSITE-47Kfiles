<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%
	if checkGIPconfig("CheckLoginPassword") and session("userID") = session("password") then
		response.write "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>"
        	response.write "<script language=javascript>alert('請變更密碼、更新負責人員姓名！');location.href='/user/userUpdate.asp';</script>"
        	response.end
	end if
 
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
   
'   ItemID = request("CtRootID")
   ItemID = session("exRootID")
   'tocURL = "Folder_toc.asp"
' 	tocURL =  "xuToc.asp"
  'if Instr(session("uGrpID")&",", "HTSD,") > 0 then 
  	tocURL =  "old_xmlToc.asp?id="&request("id")
'  	tocURL =  "xToc.asp"
'  	tocURL =  "xmlToc.asp"
  'end if
'response.write "xx=" & tocURL
'response.end
'response.write "<!--" & tocURL & "-->"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<frameset cols="<%=CatalogueFrame%>" framespacing="1" border="1" id="CatalogueMenu">
  <frame name="Catalogue" target="ForumToc" scrolling="auto" noresize src="<%=tocURL%>">
    <frame name="ForumToc" scrolling="auto" target="ShowArticle">
  <noframes>
  <body>

  <p>[,Ozs.</p>

  </body>
  </noframes>
</frameset>

</html>
