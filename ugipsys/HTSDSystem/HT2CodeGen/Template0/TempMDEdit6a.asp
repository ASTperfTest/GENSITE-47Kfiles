	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>½s­×ªí³æ</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=doneURI%>?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
