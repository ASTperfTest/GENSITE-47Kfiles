<%@ CodePage = 65001 %>
<% Response.Expires = 0
   HTProgCode = "GE1T21" %><?xml version="1.0"  encoding="utf-8" ?>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
dim CtRoot, orgRoot, dtLevel, xmlStr
leadStr = "000000000000000000000000000"

 xmlStr="<?xml version=""1.0""  encoding=""utf-8"" ?>"
 xmlStr=""
 outCode "<tocTree>"
	SQLCom = "select CatTreeRoot.* from " _
		& " CatTreeRoot RIGHT OUTER JOIN nodeinfo on cattreeroot.ctrootid = nodeinfo.ctrootid "_ 
		& " Where nodeinfo.ctrootid = '"&request("id")&"'"
	set RStree = conn.execute(SqlCom)

	while not RStree.eof
		ItemID = RStree("CtRootID")
		if session("ItemID")="" then	session("ItemID")=ItemID
		outCode "<CatTree id=""" & itemID & """ name=""" & RStree("CtRootName") & """ CtNodeKind=""C"" leadStr=""4"">"
		CtRoot = ItemID
	    traverseTree 0
	    outCode "</CatTree>"
	    RStree.MoveNext
	wend
 outCode "</tocTree>"
'    response.write xmlStr & "<HR/>"

   set oXmlReg = createObject("Microsoft.XMLDOM")
   oXmlReg.async = false
   oXmlReg.loadXML xmlStr
 if oXmlReg.parseError.reason <> "" then
  'alert "內文不符合字串比較格式!"
  'response.end
	Response.Write("oXmlReg parseError on line " &  oXmlReg.parseError.line)
	Response.Write("<BR>Reasonxx: " &  oXmlReg.parseError.reason)
	Response.End()
 end if

 for each x in oXmlReg.selectNodes("//*[CatTreeNode]")
  x.lastChild.setAttribute "leadStr", "7"
 next
 oXmlReg.selectSingleNode("tocTree").lastChild.setAttribute "leadStr", "7"


 for each x in oXmlReg.selectNodes("//CatTree")
    trTree2 x, ""
   next
'   response.write oXmlReg.xml
'   response.end


	set oxsl = server.createObject("microsoft.XMLDOM")
	oxsl.load(server.mappath("old_tocTree.xsl"))
  if oxsl.parseError.reason <> "" then 
    Response.Write("oxslhtPageDom parseError on line " &  oxsl.parseError.line)
    Response.Write("<BR>Reason: " &  oxsl.parseError.reason)
    Response.End()
  end if
	Response.ContentType = "text/HTML" 

'	response.write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbCRLF
	outString = replace(oXmlReg.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=UTF-16"">","")
	outString = replace(outString,"&amp;","&")
	response.write replace(outString,"&amp;","&")


sub trTree2(thisNode, xleadStr)
 myleadStr = xleadStr
 if myleadStr<>"" then
    if right(myleadStr,1) = "7" then
   myleadStr = left(myleadStr, len(myleadStr)-1) & "0"
  elseif right(myleadStr,1) = "4" then
   myleadStr = left(myleadStr, len(myleadStr)-1) & "1"
    end if
 end if
 myleadStr = myleadStr & thisNode.attributes.getNamedItem("leadStr").value
 thisNode.setAttribute "leadStr", myleadStr

 for each x in thisNode.selectNodes("CatTreeNode")
    trTree2 x, myleadStr
   next

end sub

sub traverseTree (parent)
 SqlCom = "SELECT * FROM CatTreeNode WHERE ctRootId = "& PkStr(CtRoot,"") _
  & " AND dataParent=" & parent & " Order by catShowOrder"
' response.write sqlCom & "<HR>"
 set RSt = conn.execute(SqlCom)

 while not RSt.eof
  outCode "<CatTreeNode id=""" & RSt("CtNodeID") & """ name="""& replace(deAmp(RSt("CatName")),"""","&quot;") _
   & """ CtNodeKind="""& RSt("CtNodeKind") _
   & """ leadStr=""4"">"
  if RSt("CtNodeKind") = "C" then   traverseTree RSt("CtNodeID")
  outCode "</CatTreeNode>" & vbCRLF
  RSt.moveNext
 wend
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
 xmlStr = xmlStr & xStr
end sub
%>