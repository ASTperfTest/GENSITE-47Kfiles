<!--#Include file = "../inc/server.inc" -->
    <html> 
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
    <title>�d�ߪ��</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
taskLable="�d��" & HTProgCap

	showHTMLHead()
	formFunction = "query"
	showForm()
	initForm()
	showHTMLTail()
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm
