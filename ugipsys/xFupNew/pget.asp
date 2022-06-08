<%@ CodePage = 65001 %>
<% 
'	Response.ContentType = "text/plain"
	Response.Expires = 0
   HTProgCode = "GC1AP5"
	 CatalogueFrame = "220,*"
	 xFile = request("xFile")
%>
<!--#include virtual = "/inc/server.inc" -->
<%
    Set fs = CreateObject("scripting.filesystemobject")
    Set xfin = fs.opentextfile(Server.MapPath(xFile))

    Do While Not xfin.AtEndOfStream
        xinStr = xfin.readline
        response.write xinStr & vbCRLF
    Loop

%>
