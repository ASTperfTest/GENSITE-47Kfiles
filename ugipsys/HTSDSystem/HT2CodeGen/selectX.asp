<%
	' http://embin.hometown.com.tw/embCodeGen/xdb1.asp?createSchema=Y&xml=電子報
	'	新聞分類資料表&formid=epNewsTyp
	
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.GetFolder(server.MapPath(".") & "\xmlSpec\")
	Set fc = f.Files
	For Each f1 in fc
		if ucase(right(Trim(f1.name), 3))="XML" then
			Set xmldoc=Server.CreateObject("Microsoft.XMLDom")
			xp = server.MapPath(".") & "\xmlSpec\"			
			xv = xmldoc.load(xp+f1.name)
%>
<%=f1.name%><-------->
<%
			set formDom = xmldoc.selectSingleNode("DataSchemaDef/dsTable/tableName")
%>
<A href="http://embin.hometown.com.tw/embCodeGen/xdb1.asp?createSchema=Y&xml=<%=left(f1.name, len(f1.name)-4)%>
&formid=<%=formDom.Text%>">
[建立]
</A>
<HR>
<%
			Set xmldoc=Nothing
			Set formDom=Nothing
		end if
	Next
	
%>