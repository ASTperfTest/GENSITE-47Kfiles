<% Response.Expires = 0
HTProgCode="HT011"
HTProgPrefix="htDField" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%
	xmlFile = request.form("xmlFile")
'response.write request.form("formSpecFile")
'response.end	
	formSpecFile = request.form("formSpecFile")

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
	
	set nxml = server.createObject("microsoft.XMLDOM")
	nxml.async = false
	nxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/formSpecNone0.xml")
	xv = nxml.load(LoadXML)
  if nxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  nxml.parseError.line)
    Response.Write("<BR>Reason: " &  nxml.parseError.reason)
    Response.End()
  end if

	progPrefix = mid(xmlFile,5)
	progPrefix = left(progPrefix,len(progPrefix)-5)
	nxml.selectSingleNode("htPage/HTProgPrefix").text = progPrefix
	nxml.selectSingleNode("htPage/HTProgCode").text = nullText(oxml.selectSingleNode("OperationContract/Code"))
	progPath = nullText(oxml.selectSingleNode("OperationContract/htCodePath"))
	nxml.selectSingleNode("htPage/HTProgPath").text = nullText(oxml.selectSingleNode("OperationContract/htCodePath"))
	nxml.selectSingleNode("htPage/HTGenPattern").text = nullText(oxml.selectSingleNode("OperationContract/htPattern"))
	
	
	nxml.selectSingleNode("//pageSpec/pageHead").text = request("pageHead")
	nxml.selectSingleNode("//pageSpec/pageFunction").text = request("pageFunction")
	
	nxml.save(Server.MapPath("/HTSDSystem/HT2CodeGen/formSpec" & progPath & "/" & formSpecFile & ".xml"))

	set nFormSpec = oxml.createElement("formSpec")
	nFormSpec.text = formSpecFile
	oxml.selectSingleNode("OperationContract").insertBefore nFormSpec, oxml.selectSingleNode("OperationContract/htPattern")
	oxml.save Server.MapPath("/HTSDSystem/UseCase/" & xmlFile & ".xml")

	response.redirect "viewContract.asp?xmlFile=" & xmlFile & ".xml"

sub processField(xfObj)
	set xnTR = xTR.cloneNode(true)
	set xnTD = xnTr.selectSingleNode("TD[@class='eTableLable']")
	
	xnTD.text = nullText(xfObj.selectSingleNode("fieldLabel")) & "："
	xnTr.selectSingleNode("//refField").text = xb& "/" & xfObj.selectSingleNode("fieldName").text
	pXml.appendChild xnTR
end sub
	
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

%>
