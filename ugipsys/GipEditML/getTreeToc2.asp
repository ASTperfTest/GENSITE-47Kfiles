<%@ CodePage = 65001 %>
<% Response.Expires = 0
Response.AddHeader "Content-Type", "text/xml; charset=utf-8"
   HTProgCode = "GE1T21" %><?xml version="1.0"  encoding="utf-8" ?>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
dim CtRoot, orgRoot, dtLevel, xmlStr
 outCode "<toc>"
	gDepth = cint(request("depth"))
  if request("id")="treeRoot" then
	SQLCom = "select * from CatTreeRoot Where pvXdmp is not null " _
		& " AND (deptID IS NULL OR deptID LIKE '" & session("deptID") & "%'" _
		& " OR '" & session("deptID") & "' LIKE deptID+'%')" _
		& " ORDER BY deptID, CtRootID"
	set RStree = conn.execute(SqlCom)
	    outCode "<tocnode id=""treeRoot"" nodeType=""C"">"
		outCode "<desc>資料上稿</desc>" 
		outCode "<contents>"

	while not RStree.eof
	    outCode "<tocnode ID=""t" & RStree("CtRootID") & """ nodeType=""C"">" 
		outCode "<desc>" & deAmp(RStree("CtRootName")) & "</desc>" 
		CtRoot = RStree("CtRootID")
		if gDepth>0 then 	    traverseTree 0, gDepth-1
	    outCode "</tocnode>"
	    RStree.MoveNext
	wend
		outCode "</contents>"
	    outCode "</tocnode>"
  elseif left(request("id"),1)="t" then
	SQLCom = "select * from CatTreeRoot Where CtRootID=" & pkStr(mid(request("id"),2),"")
	set RStree = conn.execute(SqlCom)
	if not RStree.eof then
	    outCode "<tocnode ID=""t" & RStree("CtRootID") & """ parentID=""treeRoot"" nodeType=""C"">"
		outCode "<desc>" & deAmp(RStree("CtRootName")) & "</desc>" 
		CtRoot = RStree("CtRootID")
		if gDepth>0 then 	    traverseTree 0, gDepth
	    outCode "</tocnode>"
	end if
  else
	SQLCom = "select * from CatTreeNode Where CtNodeID=" & pkStr(request("id"),"")
	set RStree = conn.execute(SqlCom)
	if not RStree.eof then
		CtRoot = RStree("CtRootID")
	    dim DataParent
	    DataParent = RStree("DataParent")
	    if DataParent = "0" then
		    outCode "<tocnode ID=""" & RStree("CtNodeID") & """ parentID=""t" _
		    	& CtRoot & """ nodeType=""" & RStree("CtNodeKind") & """>" 
	    else
		    outCode "<tocnode ID=""" & RStree("CtNodeID") & """ parentID=""" _
		    	& DataParent & """ nodeType=""" & RStree("CtNodeKind") & """>" 
	    end if
		if RStree("CtUnitID")<>"" then
			outCode "<desc><A href=""ctXMLin.asp?ItemID=" & CtRoot & "&amp;CtNodeID=" _
				& RStree("CtNodeID") & """>" & deAmp(RStree("CatName")) & "</A></desc>" 
		else
			outCode "<desc>" & deAmp(RStree("CatName")) & "</desc>" 
   		end if 
		if RStree("CtNodeKind") = "C" AND gDepth>0 then 	    traverseTree RStree("CtNodeID"), gDepth
	    outCode "</tocnode>"
	end if
  end if
 outCode "</toc>"
'    response.write xmlStr & "<HR/>"

sub traverseTree (parent, nDepth)
 SqlCom = "SELECT * FROM CatTreeNode WHERE CtRootID = "& PkStr(CtRoot,"") _
  & " AND dataParent=" & parent & " Order by CatShowOrder"
' response.write sqlCom & "<HR>"
 set RSt = conn.execute(SqlCom)
	if RSt.eof then	exit sub
	
	outCode "<contents>"
 	while not RSt.eof
	    outCode "<tocnode ID=""" & RSt("CtNodeID") & """ nodeType=""" & RSt("CtNodeKind") & """>" 
		if RSt("CtUnitID")<>"" then
			outCode "<desc><A href=""ctXMLin.asp?ItemID=" & CtRoot & "&amp;CtNodeID=" _
				& RSt("CtNodeID") & """>" & deAmp(RSt("CatName")) & "</A></desc>" 
		else
			outCode "<desc>" & deAmp(RSt("CatName")) & "</desc>" 
   		end if 
		if RSt("CtNodeKind") = "C" AND nDepth>1 then 	    traverseTree RSt("CtNodeID"), nDepth-1
	    outCode "</tocnode>"
		RSt.moveNext
	wend
	outCode "</contents>"
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
