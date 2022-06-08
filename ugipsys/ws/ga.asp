<%
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f1 = fso.OpenTextFile(server.mapPath("\ws") & "\testfile.txt", 8, true)
	
'	Response.AddHeader "Content-type", "text/html; charset=utf-8"
'	Response.ContentType = "text/html"
'	Response.Charset= "utf-8" 

'	f1.readAll
	f1.writeline("Form" & date() & ":" & time())
	f1.writeline("--form--------------"&vbCRLF)
'	f1.writeline(Request.Form.count & "=======" & vbcrlf)
'	for each y in Request.Form
'		f1.writeline(y & ":" & Request(y))
'	next
	f1.writeline("--QueryString--------------"&vbCRLF)
	for each y in Request.QueryString
		f1.writeline(y & ":" & Request(y)&vbCRLF)

		if left(request(y),12)= "xmlcontents=" then
			set oXmlDoc = Server.CreateObject("MICROSOFT.XMLDOM")
			oXmlDoc.async = false
			oXMLDoc.setProperty("ServerHTTPRequest") = true	
			f1.writeline("xml :=" & mid(request(y),13)&vbCRLF)
			v = oXMLDoc.loadXML(mid(request(y),13))
			f1.writeline("load result :=" & v &vbCRLF)
		end if 
	next


'	f1.writeline("--ServerVariables--------------"&vbCRLF)
'	for each y in Request.ServerVariables
'		f1.writeline(y & ":" & Request(y))
'	next
'	f1.writeline("--Item--------------"&vbCRLF)
'	Response.Write "haha!!" & "<HR>"

'  lngCount = Request.TotalBytes
'  f1.writeline("totalbytes="&lngCount&vbcrlf)
'  vntPostedData = Request.BinaryRead(lngCount)
'  f1.writeline("ubound="&ubound(vntPostedData)&vbcrlf)
'  for xi=0 to ubound(vntPostedData)
'	f1.writeline(xi & ")-->" & cint(vntPostedData(xi)) & vbCRLF)
'  next
		f1.writeline("---end-------------"&vbCRLF)
	f1.close
%>
<?xml version="1.0"  encoding="utf-8" ?>
<%
	response.write oXmlDoc.documentElement.xml
%>
