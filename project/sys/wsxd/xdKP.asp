<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

		response.write "<rURL>" & request("URL") & "</rURL>"

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("GipDSD") & "\xdmp" & request("mp") & ".xml"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if

  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	response.write "<MenuTitle>"&nullText(refModel.selectSingleNode("MenuTitle"))&"</MenuTitle>"
  	
  	for each xDataSet in refModel.selectNodes("DataSet")
  		processXDataSet
	next
  myTreeNode = 0
  upParent = 0

function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

%>
<!--#include file="x1Menus.inc" -->
</hpMain>
