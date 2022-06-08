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
	LoadXML = server.mappath("/HTSDSystem/formSpecBList0.xml")
	xv = nxml.load(LoadXML)
  if nxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  nxml.parseError.line)
    Response.Write("<BR>Reason: " &  nxml.parseError.reason)
    Response.End()
  end if

	progPrefix = mid(xmlFile,5)
	progPrefix = left(progPrefix,len(progPrefix)-4)
	nxml.selectSingleNode("htPage/HTProgPrefix").text = progPrefix
	nxml.selectSingleNode("htPage/HTProgCode").text = nullText(oxml.selectSingleNode("OperationContract/Code"))
	progPath = nullText(oxml.selectSingleNode("OperationContract/htCodePath"))
	nxml.selectSingleNode("htPage/HTProgPath").text = nullText(oxml.selectSingleNode("OperationContract/htCodePath"))
	nxml.selectSingleNode("htPage/HTGenPattern").text = nullText(oxml.selectSingleNode("OperationContract/htPattern"))
	
	
	joinCount = 0
	xSelectList = ""
	for each ex in request.form
'		response.write ex & "==>" & request.form(ex) & "<BR>"
		if right(ex,12) = "_pickedField" then
			xTable = left(ex, len(ex)-12)
			response.write xTable & "<HR>"
			set xTableNode = nxml.selectSingleNode("//resultSet/fieldList").cloneNode(true)
			set xDetailNode = nxml.selectSingleNode("//detailRow")
			xTableNode.selectSingleNode("tableName").text = xTable			
			if request("masterTable") = xTable then
				set nFormSpec = nxml.createElement("masterTable")
				nFormSpec.text = "Y"
				xTableNode.insertBefore nFormSpec, xTableNode.selectSingleNode("field")
				tbPrefix = "htx."
				nxml.selectSingleNode("//resultSet/sql/fromList").text = xTable & " AS htx"
			else
				joinCount = joinCount + 1
				tbPrefix = "htj" & joinCount & "."
				set nFormSpec = nxml.selectSingleNode("//resultSet/fieldListLink/fkLink").cloneNode(true)
				nFormSpec.selectSingleNode("asAlias").text = "htj" & joinCount
				xTableNode.insertBefore nFormSpec, xTableNode.selectSingleNode("field")
			end if
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
'				set xFieldNode = nxml.selectSingleNode("//resultSet/fieldList/field").cloneNode(true)
'				xFieldNode.selectSingleNode("fieldName").text = xFldArray(xfi)
				set xFieldNode = fxml.selectSingleNode( _
					"//dsTable[tableName='" & xTable & "']/fieldList/field[fieldName='" & xFldArray(xfi) & "']")

				set xparamKind = nxml.selectSingleNode("//resultSet/valueType").cloneNode(true)
'				xFieldNode.appendChild xparamKind
				xFieldNode.insertBefore xparamKind, xFieldNode.selectSingleNode("fieldName")

				xTableNode.appendChild xFieldNode
				xSelectList = xSelectList & ", " & tbPrefix & xFldArray(xfi)
				set xColNode = nxml.selectSingleNode("//detailRow/colSpec").cloneNode(true)
				xColNode.selectSingleNode("colLabel").text = xFieldNode.selectSingleNode("fieldLabel").text
				xColNode.selectSingleNode("content/refField").text = xFldArray(xfi)
				xDetailNode.appendChild xColNode
			next

			xTableNode.removeChild xTableNode.selectSingleNode("field")

			for each xpk in fxml.selectNodes("//dsTable[tableName='" & xTable & "']/fieldList/field[isPrimaryKey='Y']")
				if nullText(xTableNode.selectSingleNode("field[fieldName='" _
					& xpk.selectSingleNode("fieldName").text & "']")) = "" then
					xTableNode.appendChild xpk
				end if
			next

			nxml.selectSingleNode("//resultSet").insertBefore _
				xTableNode, nxml.selectSingleNode("//resultSet/fieldListLink")
		end if
	next
	nxml.selectSingleNode("//resultSet").removeChild nxml.selectSingleNode("//resultSet/fieldList")
	nxml.selectSingleNode("//resultSet").removeChild nxml.selectSingleNode("//resultSet/fieldListLink")
	nxml.selectSingleNode("//resultSet").removeChild nxml.selectSingleNode("//resultSet/valueType")
	nxml.selectSingleNode("//detailRow").removeChild nxml.selectSingleNode("//detailRow/colSpec")

	nxml.selectSingleNode("//resultSet/sql/selectList").text = mid(xSelectList,3)
	
'	nxml.selectSingleNode("//htForm/formModel").setAttribute "id", mid(xmlFile,5)
'	nxml.selectSingleNode("//pageSpec/formUI").setAttribute "ref", mid(xmlFile,5)
	nxml.selectSingleNode("//pageSpec/pageHead").text = request("pageHead")
	nxml.selectSingleNode("//pageSpec/pageFunction").text = request("pageFunction")
	
	set xov1AnchorList = oxml.selectSingleNode("//AnchorList")
	set nxAnchorList = nxml.selectSingleNode("//aidLinkList")
			response.write "Hello4<HR>"	
	for each xAnchor in xov1AnchorList.selectNodes("Anchor")
			response.write nullText(xAnchor.selectSingleNode("AnchorType")) & "Hello5<HR>"
		if left(xAnchor.selectSingleNode("AnchorType").text,4) <> "Link" then		
			response.write "Hello6<HR>"			
			nxAnchorList.appendChild xAnchor.cloneNode(true)
		end if
	next
	
'	set pXml = nxml.selectSingleNode("//pxHTML/CENTER/TABLE")
'	set xTR = pXml.selectSingleNode("TR")
'	set xFont = xTR.selectSingleNode("TD[@class='lightbluetable']/font")
'	for each xfList in nxml.selectNodes("//formModel/fieldList")
'		xb = nullText(xfList.selectSingleNode("tableName"))
'		for each xf in xfList.selectNodes("field")
'			processField xf
'		next
'	next
'	pxml.removeChild xTR

	nxml.save(Server.MapPath("/HTSDSystem/HT2CodeGen/formSpec" & progPath & "/" & formSpecFile & ".xml"))

	set nFormSpec = oxml.createElement("formSpec")
	nFormSpec.text = formSpecFile
	oxml.selectSingleNode("OperationContract").insertBefore nFormSpec, oxml.selectSingleNode("OperationContract/htPattern")
	oxml.save Server.MapPath("/HTSDSystem/UseCase/" & xmlFile & ".xml")

	response.redirect "viewContract.asp?xmlFile=" & xmlFile & ".xml"


sub processField(xfObj)
	set xnTR = xTR.cloneNode(true)
	set xnTD = xnTr.selectSingleNode("TD[@class='lightbluetable']")
	
	xnTD.text = nullText(xfObj.selectSingleNode("fieldLabel")) & "："
	if nullText(xfObj.selectSingleNode("canNull")) = "N" then _
		xnTD.insertBefore xFont.cloneNode(true), xnTD.childNodes.item(0)
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
