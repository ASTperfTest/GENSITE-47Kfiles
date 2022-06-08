<% @ CodePage = 65001 %>
<%
'// purpose: ?? emily ??download engmoe ?any type of files from server.
'// date: 2006/7/13
'// test

Response.Buffer = True
Response.Clear

HTProgCode = "GC1AP5"
CatalogueFrame = "220,*"
%>
<!--#include virtual = "/inc/server.inc" -->
<%
dim xFile
dim downFile
dim s
dim FSO
dim f

xFile = request("xFile") 
DownFile = Server.MapPath(xFile)

Set s = Server.CreateObject("ADODB.Stream")
s.Open
s.Type = 1

Set FSO = Server.CreateObject("Scripting.FileSystemObject")

if not FSO.FileExists(DownFile) then
    Response.Write("<h1>Error:</h1>" & DownFile & " does not exist<p>")
    Response.End
end if

Set f = fso.GetFile(DownFile)
intFilelength = f.size
s.LoadFromFile(DownFile)

Response.AddHeader "Content-Disposition", "attachment; filename=" & f.name
Response.AddHeader "Content-Length", intFilelength
Response.CharSet = "UTF-8"
Response.ContentType = "application/octet-stream"
Response.BinaryWrite s.Read
Response.Flush

s.Close
Set s = Nothing
Set FSO = Nothing 

%>