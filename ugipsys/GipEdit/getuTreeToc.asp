<%@ CodePage = 65001 %>
<% Response.Expires = 0
Response.AddHeader "Content-Type", "text/xml; charset=utf-8"
   HTProgCode = "GE1T21" %><?xml version="1.0"  encoding="utf-8" ?>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
dim CtRoot, orgRoot, dtLevel, xmlStr
 outCode "<root>"
	gDepth = cint(request("depth"))
  if request("id")="treeRoot" then
	SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
		& " AND (deptID IS NULL OR deptID LIKE '" & session("deptID") & "%'" _
		& " OR '" & session("deptID") & "' LIKE deptID+'%')" _
		& " ORDER BY deptID, CtRootID"
	set RStree = conn.execute(SqlCom)
	    outCode "<node>"
		outCode "<ID>treeRoot</ID>" 
		outCode "<nodeType>C</nodeType>" 
		outCode "<desc>資料上稿</desc>" 
		outCode "<children>"

	while not RStree.eof
	  CtRoot = RStree("CtRootID")
	  xSql = "SELECT count(*) FROM CtUserSet AS u JOIN CatTreeNode AS t ON u.CtNodeID=t.CtNodeID" _
		& " AND u.UserID=" & pkStr(session("userID"),"") _
		& " WHERE CtRootID = "& PkStr(CtRoot,"")
	  set RS= conn.execute(xSql)
	  if RS(0) > 0 then	
	    outCode "<node>"
		outCode "<sql>t" & xSql & "</sql>" 
		outCode "<ID>t" & RStree("CtRootID") & "</ID>" 
		outCode "<nodeType>C</nodeType>" 
		outCode "<desc>" & deAmp(RStree("CtRootName")) & "</desc>" 
		if gDepth>0 then 	    traverseTree 0, gDepth-1
	    outCode "</node>"
	  end if
	    RStree.MoveNext
	wend
		outCode "</children>"
	    outCode "</node>"
  elseif left(request("id"),1)="t" then
	SQLCom = "select * from CatTreeRoot Where CtRootID=" & pkStr(mid(request("id"),2),"")
	set RStree = conn.execute(SqlCom)
	if not RStree.eof then
	    outCode "<node>"
		outCode "<ID>t" & RStree("CtRootID") & "</ID>" 
		outCode "<nodeType>C</nodeType>" 
		outCode "<desc>" & deAmp(RStree("CtRootName")) & "</desc>" 
		CtRoot = RStree("CtRootID")
		if gDepth>0 then 	    traverseTree 0, gDepth
	    outCode "</node>"
	end if
  else
	SQLCom = "select * from CatTreeNode Where CtNodeID=" & pkStr(request("id"),"")
	set RStree = conn.execute(SqlCom)
	if not RStree.eof then
		CtRoot = RStree("CtRootID")
	    outCode "<node>"
		outCode "<ID>" & RStree("CtNodeID") & "</ID>" 
		outCode "<nodeType>" & RStree("CtNodeKind") & "</nodeType>" 
		if RStree("CtUnitID")<>"" then
			outCode "<desc><A href=""ctXMLin.asp?ItemID=" & CtRoot & "&amp;CtNodeID=" _
				& RStree("CtNodeID") & """>" & deAmp(RStree("CatName")) & "</A></desc>" 
		else
			outCode "<desc>" & deAmp(RStree("CatName")) & "</desc>" 
   		end if 
		if RStree("CtNodeKind") = "C" AND gDepth>0 then 	    traverseTree RStree("CtNodeID"), gDepth
	    outCode "</node>"
	end if
  end if
 outCode "</root>"
'    response.write xmlStr & "<HR/>"

sub traverseTree (parent, nDepth)
 SqlCom = "SELECT * FROM CtUserSet AS u JOIN CatTreeNode AS t ON u.CtNodeID=t.CtNodeID" _
  & " WHERE t.CtRootID = "& PkStr(CtRoot,"") _
  & " AND u.UserID=" & pkStr(session("userID"),"") _
  & " AND t.dataParent=" & parent & " Order by CatShowOrder"
' response.write sqlCom & "<HR>"
 set RSt = conn.execute(SqlCom)
	if RSt.eof then	exit sub
	
	outCode "<children>"
 	while not RSt.eof
	    outCode "<node>"
		outCode "<ID>" & RSt("CtNodeID") & "</ID>" 
		outCode "<nodeType>" & RSt("CtNodeKind") & "</nodeType>" 
		if RSt("CtUnitID")<>"" then
			outCode "<desc><A href=""ctXMLin.asp?ItemID=" & CtRoot & "&amp;CtNodeID=" _
				& RSt("CtNodeID") & """>" & deAmp(RSt("CatName")) & "</A></desc>" 
		else
			outCode "<desc>" & deAmp(RSt("CatName")) & "</desc>" 
   		end if 
		if RSt("CtNodeKind") = "C" AND nDepth>1 then 	    traverseTree RSt("CtNodeID"), nDepth-1
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
