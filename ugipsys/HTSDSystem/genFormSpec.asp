<% Response.Expires = 0
HTProgCode="HT011"
HTProgPrefix="htDField" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%
	xmlFile = request.form("xmlFile")
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
	LoadXML = server.mappath("/HTSDSystem/formSpec0.xml")
	xv = nxml.load(LoadXML)
  if nxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  nxml.parseError.line)
    Response.Write("<BR>Reason: " &  nxml.parseError.reason)
    Response.End()
  end if

	nxml.selectSingleNode("htPage/HTProgPrefix").text = mid(xmlFile,5)
	nxml.selectSingleNode("htPage/HTProgCode").text = nullText(oxml.selectSingleNode("OperationContract/Code"))
	nxml.selectSingleNode("htPage/HTGenPattern").text = nullText(oxml.selectSingleNode("OperationContract/htPattern"))
	
	
	for each ex in request.form
'		response.write ex & "==>" & request.form(ex) & "<BR>"
		if right(ex,12) = "_pickedField" then
			xTable = left(ex, len(ex)-12)
			response.write xTable & "<HR>"
			set xTableNode = nxml.selectSingleNode("//formModel/fieldList").cloneNode(true)
			xTableNode.selectSingleNode("tableName").text = xTable

			set fxml = server.createObject("microsoft.XMLDOM")
			fxml.async = false
			fxml.setProperty "ServerHTTPRequest", true
			LoadXML = server.mappath("/HTSDSystem/xmlSpec/" & xTable & ".xml")
			xv = fxml.load(LoadXML)
  			if fxml.parseError.reason <> "" then
    			Response.Write(xTable & ".xml parseError on line " &  fxml.parseError.line)
    			Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    			Response.End()
  			end if

			xFldArray = split(request(ex),", ")
			for xfi = 0 to uBound(xFldArray)
				response.write "--" & xFldArray(xfi) & "<BR>"
'				set xFieldNode = nxml.selectSingleNode("//formModel/fieldList/field").cloneNode(true)
'				xFieldNode.selectSingleNode("fieldName").text = xFldArray(xfi)
				set xFieldNode = fxml.selectSingleNode( _
					"//dsTable[tableName='" & xTable & "']/fieldList/field[fieldName='" & xFldArray(xfi) & "']")
				xTableNode.appendChild xFieldNode
			next

			xTableNode.removeChild xTableNode.selectSingleNode("field")
			nxml.selectSingleNode("//formModel").insertBefore _
				xTableNode, nxml.selectSingleNode("//formModel/anchorList")
		end if
	next
	nxml.selectSingleNode("//formModel").removeChild nxml.selectSingleNode("//formModel/fieldList")
	
	nxml.selectSingleNode("//htForm/formModel").setAttribute "id", mid(xmlFile,5)
	nxml.selectSingleNode("//pageSpec/formUI").setAttribute "ref", mid(xmlFile,5)
	nxml.selectSingleNode("//pageSpec/pageHead").text = request("pageHead")
	nxml.selectSingleNode("//pageSpec/pageFunction").text = request("pageFunction")
	
	set pXml = nxml.selectSingleNode("//pxHTML/CENTER/TABLE")
	set xTR = pXml.selectSingleNode("TR")
	set xFont = xTR.selectSingleNode("TD[@class='lightbluetable']/font")
	for each xfList in nxml.selectNodes("//formModel/fieldList")
		xb = nullText(xfList.selectSingleNode("tableName"))
		for each xf in xfList.selectNodes("field")
			processField xf
		next
	next
	pxml.removeChild xTR

sub processField(xfObj)
	set xnTR = xTR.cloneNode(true)
	set xnTD = xnTr.selectSingleNode("TD[@class='lightbluetable']")
	
	xnTD.text = nullText(xfObj.selectSingleNode("fieldLabel")) & "："
	if nullText(xfObj.selectSingleNode("canNull")) = "N" then _
		xnTD.insertBefore xFont.cloneNode(true), xnTD.childNodes.item(0)
	xnTr.selectSingleNode("//refField").text = xb& "/" & xfObj.selectSingleNode("fieldName").text
	pXml.appendChild xnTR
end sub


	
	nxml.save(Server.MapPath("/HTSDSystem/HT2CodeGen/formSpec/" & formSpecFile & ".xml"))

	set nFormSpec = oxml.createElement("formSpec")
	nFormSpec.text = formSpecFile
	oxml.selectSingleNode("OperationContract").insertBefore nFormSpec, oxml.selectSingleNode("OperationContract/htPattern")
	oxml.save Server.MapPath("/HTSDSystem/UseCase/" & xmlFile & ".xml")

	response.redirect "viewContract.asp?xmlFile=" & xmlFile & ".xml"


function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

%>
