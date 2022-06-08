<% Response.Expires = 0
HTProgCode="HT011"
HTProgPrefix="htDField" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<!--#include virtual = "/inc/xdbutil.inc" -->
<%
	xmlFile = request.queryString("xmlFile")
	
	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
	LoadXML = server.mappath("/HTSDSystem/UseCase/" & xmlFile & ".xml")
	xv = oxml.load(LoadXML)
  if oxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  oxml.parseError.line)
    Response.Write("<BR>Reason: " &  oxml.parseError.reason)
    Response.End()
  end if

	htSpecPattern = nullText(oxml.selectSingleNode("OperationContract/htSpecPattern"))
	set vo = oxml.selectSingleNode("//Version")
  if nullText(vo) = "" then
  	response.write "version not exist"
	response.end
  end if
%>
<BODY>
<FORM method="POST" action="genFormSpec<%=htSpecPattern%>.asp">
<table border=0 width="100%">
<tr><td>
CodePath: <INPUT name="codePath" size="50" value="<%=Server.MapPath(nullText(oxml.selectSingleNode("OperationContract/htCodePath")))%>" readonly><BR>
<INPUT type="hidden" name="xmlFile" value="<%=request.queryString("xmlFile")%>">
formSpec 檔名：<INPUT name="formSpecFile" value="<%=mid(xmlFile,5)%>"><BR>
pageHead：<INPUT name="pageHead" value=""><BR>
pageFunction<INPUT name="pageFunction" value="">
<td align="right">
<INPUT type="submit" value="確定">
<INPUT type="reset">
</td></tr></table>
<HR>
<%
	for each xEntity in vo.selectNodes("ContractSpec/htDEntityList/htDEntity")
%>
-<INPUT TYPE="radio" NAME="masterTable" VALUE="<%=xEntity.text%>">(Master)
<span onClick="VBS: expEntity 'xd_<%=xEntity.text%>'" style="cursor:hand;"><%=xEntity.text%>：</span>
<BUTTON onClick="VBS: checkAll '<%=xEntity.text%>_pickedField'">All</BUTTON>&nbsp;
<BUTTON onClick="VBS: uncheckAll '<%=xEntity.text%>_pickedField'">None</BUTTON>
-<INPUT TYPE="radio" NAME="detailTable" VALUE="<%=xEntity.text%>">(Detail)
<DIV id="xd_<%=xEntity.text%>" style="margin-left:50px; display=block;">
<%
		fSql = "SELECT f.entityID, f.xfieldName, f.xfieldLabel, e.tableName " _
			& " FROM (htDField AS f JOIN htDentity AS e ON e.entityID = f.entityID " _
			& " AND e.tableName=" & pkStr(xEntity.text,")") _
			& " ORDER BY f.xFieldSeq"
		set RSlist = conn.execute(fSql)
		while not RSlist.eof
'			response.write RSlist("tableName") & "/" & RSlist("xFieldName") & "(" & RSlist("xFieldLabel") & ")<BR>"
%>
<INPUT TYPE="checkbox" NAME="<%=RSlist("tableName")%>_pickedField" VALUE="<%=RSlist("xFieldName")%>">
<%
'			response.write RSlist("tableName") & " / " 
			response.write RSlist("xFieldName") & " (" & RSlist("xFieldLabel") & ")<BR>"
			RSlist.moveNext
		wend
		response.write "</DIV><HR>"
	next
%>
</FORM>
<script language=vbs>
sub expEntity (xdivName)
	set oDiv = document.all(xdivName)
	if oDiv.style.display="none" then
		oDiv.style.display="block" 
	else
		oDiv.style.display="none" 
	end if
end sub

sub checkAll (xName)
	for each xc in document.all(xName)
		xc.checked = true
	next
end sub

sub uncheckAll (xName)
	for each xc in document.all(xName)
		xc.checked = false
	next
end sub
</script>
</BODY>
</HTML>
<%

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

%>
