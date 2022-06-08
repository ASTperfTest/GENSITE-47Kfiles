<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<!--#Include file = "time.inc" -->
<!--#Include virtual = "/inc/GSSUtility.inc" -->
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function


function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

'	response.write "**" & request("xdURL") & "**"
'	response.end
'if request("gstyle") <> "" then
'	session("gstyle")=request("gstyle")
'end if

	if request("memID") <> "" Then
		Session("memID") = Request("memID")
	end if
	'從年網帶回寫好的HTML與SCRIPT
	testXml = session("harvest_myXDURL") & "/xdhqp.asp?mp=4&CtNode=1487&CtUnit=261&BaseDSD=35"
	response.write "<pHTML><![CDATA["
	
		set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
		xv = oxml.load(testXml)
		response.write oxml.selectSingleNode("pHTML").text	

	response.write "]]></pHTML>"

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
'		response.write LoadXML & "<HR>"
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
	response.write "<myStyle>"&nullText(refModel.selectSingleNode("MpStyle"))&"</myStyle>"
	response.write "<mp>"&request("mp")&"</mp>"
	
	'-----會員登入區塊-----GSSUtility
	call GetLoginInfo(request("memID"), request("gstyle")) 
	
'	response.write refModel.xml
  myTreeNode = 0
  upParent = 0
   for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
	processXDataSet
  next

%>
<!--#include file="x1Menus.inc" -->
</hpMain>
