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
	qStr = request.queryString
	if instr(qStr, "mp=1") <> 0 Then
		qStr = replace(qStr,"mp=1;","mp=" & session("harvest_mpTree"))
	end if
	testXml = session("harvest_myXDURL") & "/xdhlp.asp?" & qStr
	response.write "<pHTML><![CDATA["
	
		set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true
	xv = oxml.load(testXml)
	if oxml.parseError.reason <> "" then 
    		Response.Write("oxml parseError on line " &  oxml.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  oxml.parseError.reason)
    		Response.End()
  		end if
	set oxsl = server.createObject("microsoft.XMLDOM")	
	oxsl.async = false
	oxsl.setProperty("ServerHTTPRequest") = true
	'從豐年網讀取xsl將搜尋結果寫成html
	LoadXSL = session("harvest_url") & "xslGip/"& session("harvest_myStyle") &"/hlp.xsl"
	oxsl.load(LoadXSL)
	outString = replace(oxml.transformNode(oxsl),"<META http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">","")
	outString = replace(outString," xmlns=""""","")
	'response.write replace(outString,"&amp;","&")
	'response.write oxml.transformNode(oxsl)
	response.write replace(outString,"&amp;","&")
	
	response.write "]]></pHTML>"

		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
		'response.write LoadXML & "<HR>"
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
