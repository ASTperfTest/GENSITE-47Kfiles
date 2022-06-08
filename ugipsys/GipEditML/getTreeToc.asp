<%@ CodePage = 65001 %>
<% Response.Expires = 0
Response.AddHeader "Content-Type", "text/xml; charset=utf-8"
%><?xml version="1.0"  encoding="utf-8" ?>
<%
	HTProgCode = "GE1T21"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
dim CtRoot, orgRoot, dtLevel, xmlStr
 outCode "<root>"
	gDepth = cint(request("depth"))
  if request("id")="treeRoot" then
	SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
		& " AND (deptId IS NULL OR deptId LIKE '" & session("deptId") & "%'" _
		& " OR '" & session("deptId") & "' LIKE deptId+'%')" _
		& " ORDER BY deptId, ctRootId"
	set RStree = conn.execute(SqlCom)
	    outCode "<node>"
		outCode "<ID>treeRoot</ID>" 
		outCode "<nodeType>C</nodeType>" 
		outCode "<desc>資料上稿</desc>" 
		outCode "<children>"

	while not RStree.eof
	    outCode "<node>"
		outCode "<ID>t" & RStree("ctRootId") & "</ID>" 
		outCode "<nodeType>C</nodeType>" 
		outCode "<desc>" & deAmp(RStree("CtRootName")) & "</desc>" 
		CtRoot = RStree("ctRootId")
		if gDepth>0 then 	    traverseTree 0, gDepth-1
	    outCode "</node>"
	    RStree.MoveNext
	wend
		outCode "</children>"
	    outCode "</node>"
  elseif left(request("id"),1)="t" then
	SQLCom = "select * from CatTreeRoot Where ctRootId=" & pkStr(mid(request("id"),2),"")
	set RStree = conn.execute(SqlCom)
	if not RStree.eof then
	    outCode "<node>"
		outCode "<ID>t" & RStree("ctRootId") & "</ID>" 
		outCode "<nodeType>C</nodeType>" 
		outCode "<desc>" & deAmp(RStree("CtRootName")) & "</desc>" 
		CtRoot = RStree("ctRootId")
		if gDepth>0 then 	    traverseTree 0, gDepth
	    outCode "</node>"
	end if
  else
	SQLCom = "select * from CatTreeNode Where ctNodeId=" & pkStr(request("id"),"")
	set RStree = conn.execute(SqlCom)
	if not RStree.eof then
		CtRoot = RStree("ctRootId")
	    outCode "<node>"
		outCode "<ID>" & RStree("ctNodeId") & "</ID>" 
		outCode "<nodeType>" & RStree("CtNodeKind") & "</nodeType>" 
		if RStree("CtUnitID")<>"" then
			outCode "<desc><A href=""ctXMLin.asp?ItemID=" & CtRoot & "&amp;ctNodeId=" _
				& RStree("ctNodeId") & """>" & deAmp(RStree("CatName")) & "</A></desc>" 
		else
			outCode "<desc>" & deAmp(RStree("CatName")) & "</desc>" 
   		end if 
		if RStree("CtNodeKind") = "C" AND gDepth>0 then 	    traverseTree RStree("ctNodeId"), gDepth
	    outCode "</node>"
	end if
  end if
 outCode "</root>"
'    response.write xmlStr & "<HR/>"

sub traverseTree (parent, nDepth)
 SqlCom = "SELECT * FROM CatTreeNode WHERE ctRootId = "& PkStr(CtRoot,"") _
  & " AND dataParent=" & parent & " Order by catShowOrder"
' response.write sqlCom & "<HR>"
 set RSt = conn.execute(SqlCom)
	if RSt.eof then	exit sub
	
	outCode "<children>"
 	while not RSt.eof
	    outCode "<node>"
		outCode "<ID>" & RSt("ctNodeId") & "</ID>" 
		outCode "<nodeType>" & RSt("CtNodeKind") & "</nodeType>" 
		if RSt("CtUnitID")<>"" then
			outCode "<desc><A href=""ctXMLin.asp?ItemID=" & CtRoot & "&amp;ctNodeId=" _
				& RSt("ctNodeId") & """>" & deAmp(RSt("CatName")) & "</A></desc>" 
		else
			outCode "<desc>" & deAmp(RSt("CatName")) & "</desc>" 
   		end if 
		if RSt("CtNodeKind") = "C" AND nDepth>1 then 	    traverseTree RSt("ctNodeId"), nDepth-1
	    outCode "</node>"
		RSt.moveNext
	wend
	outCode "</children>"
end sub

function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
   deAmp=""
   exit function
  end if
   deAmp = replace(xs,"&","&amp;")
end function

sub outCode(xstr)
' xmlStr = xmlStr & xStr
	response.write xstr & vbCRLF
end sub
%>