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
	LoadXML = server.mappath("/HTSDSystem/formSpecMDList0.xml")
	xv = nxml.load(LoadXML)
  if nxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  nxml.parseError.line)
    Response.Write("<BR>Reason: " &  nxml.parseError.reason)
  end if

	progPrefix = mid(xmlFile,5)
	if lCase(right(progPrefix,4)) = "list" then
		progPrefix = left(progPrefix,len(progPrefix)-4)
	elseif lCase(right(progPrefix,8)) = "editlist" then
		progPrefix = left(progPrefix,len(progPrefix)-8)
	end if
	nxml.selectSingleNode("htPage/HTProgPrefix").text = progPrefix
	nxml.selectSingleNode("htPage/HTProgCode").text = nullText(oxml.selectSingleNode("OperationContract/Code"))
	progPath = nullText(oxml.selectSingleNode("OperationContract/htCodePath"))
	nxml.selectSingleNode("htPage/HTProgPath").text = nullText(oxml.selectSingleNode("OperationContract/htCodePath"))
	nxml.selectSingleNode("htPage/HTGenPattern").text = nullText(oxml.selectSingleNode("OperationContract/htPattern"))

	set xDetailNode = nxml.selectSingleNode("//detailRow")
	joinCount = 0
	for each ex in request.form
'		response.write ex & "==>" & request.form(ex) & "<BR>"
		if right(ex,12) = "_pickedField" then
			xTable = left(ex, len(ex)-12)
			response.write xTable & "<HR>"
			set xTableNode = nxml.selectSingleNode("//formModel/fieldList").cloneNode(true)
			xTableNode.selectSingleNode("tableName").text = xTable
			if request("masterTable") = xTable then
				set nFormSpec = nxml.createElement("masterTable")
				nFormSpec.text = "Y"
				xTableNode.insertBefore nFormSpec, xTableNode.selectSingleNode("field")
				tbPrefix = "htx."
			elseif request("detailTable") = xTable then
				nxml.selectSingleNode("//detailRow/detailRef").text = xTable
				tbPrefix = "dhtx."
				set nFormSpec = nxml.selectSingleNode("//formModel/fkLink").cloneNode(true)
				nFormSpec.selectSingleNode("asAlias").text = "dhtx"
				xTableNode.insertBefore nFormSpec, xTableNode.selectSingleNode("field")
			else
				joinCount = joinCount + 1
				tbPrefix = "htj" & joinCount & "."
				set nFormSpec = nxml.selectSingleNode("//formModel/fkLink").cloneNode(true)
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
'				set xFieldNode = nxml.selectSingleNode("//formModel/fieldList/field").cloneNode(true)
'				xFieldNode.selectSingleNode("fieldName").text = xFldArray(xfi)
				set xFieldNode = fxml.selectSingleNode( _
					"//dsTable[tableName='" & xTable & "']/fieldList/field[fieldName='" & xFldArray(xfi) & "']")
				xTableNode.appendChild xFieldNode
				if request("masterTable") <> xTable then
					set xColNode = nxml.selectSingleNode("//detailRow/colSpec").cloneNode(true)
					xColNode.selectSingleNode("colLabel").text = xFieldNode.selectSingleNode("fieldLabel").text
					xColNode.selectSingleNode("content/refField").text = xTable & "/" & xFldArray(xfi)
					xDetailNode.appendChild xColNode
				end if
			next

			xTableNode.removeChild xTableNode.selectSingleNode("field")

			for each xpk in fxml.selectNodes("//dsTable[tableName='" & xTable & "']/fieldList/field[isPrimaryKey='Y']")
				if nullText(xTableNode.selectSingleNode("field[fieldName='" _
					& xpk.selectSingleNode("fieldName").text & "']")) = "" then
					xTableNode.appendChild xpk
				end if
			next
			
			nxml.selectSingleNode("//formModel").insertBefore _
				xTableNode, nxml.selectSingleNode("//formModel/anchorList")
		end if
	next
	nxml.selectSingleNode("//formModel").removeChild nxml.selectSingleNode("//formModel/fieldList")
	nxml.selectSingleNode("//formModel").removeChild nxml.selectSingleNode("//formModel/fkLink")
	nxml.selectSingleNode("//detailRow").removeChild nxml.selectSingleNode("//detailRow/colSpec")
	
	nxml.selectSingleNode("//htForm/formModel").setAttribute "id", mid(xmlFile,5)
	nxml.selectSingleNode("//pageSpec/formUI").setAttribute "ref", mid(xmlFile,5)
	nxml.selectSingleNode("//pageSpec/pageHead").text = request("pageHead")
	nxml.selectSingleNode("//pageSpec/pageFunction").text = request("pageFunction")
	nxml.selectSingleNode("//pageSpec/formUI/masterRef").text = request("masterTable")
	
	set xov1AnchorList = oxml.selectSingleNode("//AnchorList")
	set nxAnchorList = nxml.selectSingleNode("//aidLinkList")
	for each xAnchor in xov1AnchorList.selectNodes("Anchor")
		response.write xAnchor.selectSingleNode("AnchorLabel").text & "<HR>"
		if left(xAnchor.selectSingleNode("AnchorType").text,4) <> "Link" then		
			nxAnchorList.appendChild xAnchor.cloneNode(true)
		end if
	next

	set pXml = nxml.selectSingleNode("//pxHTML/CENTER/TABLE")
	set xTR = pXml.selectSingleNode("TR")
	set xFont = xTR.selectSingleNode("TD[@class='eTableLable']/font")
	for each xfList in nxml.selectNodes("//formModel/fieldList[masterTable='Y']")
		xb = nullText(xfList.selectSingleNode("tableName"))
		for each xf in xfList.selectNodes("field[inputType!='hidden']")
			processField xf
		next
	next
	pxml.removeChild xTR
	set xTR = pXml.selectSingleNode("TR/TD[@class='eTableContent']")
	for each xfList in nxml.selectNodes("//formModel/fieldList[masterTable='Y']")
		xb = nullText(xfList.selectSingleNode("tableName"))
		for each xf in xfList.selectNodes("field[inputType='hidden']")
			set xnTF = xTR.selectSingleNode("refField").cloneNode(true)
			xnTF.text = xb & "/" & xf.selectSingleNode("fieldName").text
			xTR.appendChild xnTF
		next
	next

sub processField(xfObj)
	set xnTR = xTR.cloneNode(true)
	set xnTD = xnTr.selectSingleNode("TD[@class='eTableLable']")
	
	xnTD.text = nullText(xfObj.selectSingleNode("fieldLabel")) & "："
	if nullText(xfObj.selectSingleNode("canNull")) = "N" then _
		xnTD.insertBefore xFont.cloneNode(true), xnTD.childNodes.item(0)
	xnTr.selectSingleNode("//refField").text = xb& "/" & xfObj.selectSingleNode("fieldName").text
	pXml.appendChild xnTR
end sub

	nxml.save(Server.MapPath("/HTSDSystem/HT2CodeGen/formSpec" & progPath & "/" & formSpecFile & ".xml"))

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
