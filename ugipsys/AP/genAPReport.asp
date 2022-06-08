<%@ CodePage = 65001 %>
<% HTProgCap="AP"
HTProgCode="HT003"
HTProgPrefix="AP" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%

	set nxml = server.createObject("microsoft.XMLDOM")
	nxml.async = false
	nxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/APlist0.xml")
	xv = nxml.load(LoadXML)
  if nxml.parseError.reason <> "" then
    Response.Write("XML parseError on line " &  nxml.parseError.line)
    Response.Write("<BR>Reason: " &  nxml.parseError.reason)
    Response.End()
  end if
  	set oCatList = nxml.selectSingleNode("CatList")
  	set oUseCase = nxml.selectSingleNode("//UseCase")


	fSql = "SELECT AP.*,APCat.APCatCName FROM AP Inner Join APCat ON AP.APCat=APCat.APCatID" _
		& " WHERE AP.APpath like N'/HTSDSystem/%'" _
		& " ORDER BY APCat.APseq, AP.APorder"

	set RS = conn.execute(fSql)
	
	response.write "<TABLE BORDER><TR>"
	for i = 0 to RS.fields.count-1
		response.write "<TH>" & RS.fields(i).name
	next

	xCatID = ""
	while not RS.eof
		if RS("APcat") <> xCatID then
			if xCatID <> "" then _
				oUseCaseList.removeChild oUseCaseList.selectSingleNode("UseCase")
			xCatID = RS("APcat")
			set nAPCat = oCatList.selectSingleNode("APCat").cloneNode(true)
			nAPCat.selectSingleNode("CatID").text = xCatID
			nAPCat.selectSingleNode("CatName").text = RS("APCatCName")
			oCatList.appendChild nApCat
			
			set oUseCaseList = oCatList.lastChild.selectSingleNode("UseCaseList")
		end if
		response.write "<TR>"
		for i = 0 to RS.fields.count-1
			response.write "<TD>" & RS.fields(i)
		next
		fillUseCase
	
		RS.moveNext
	wend
	response.write "</TABLE>"

	if xCatID <> "" then _
		oUseCaseList.removeChild oUseCaseList.selectSingleNode("UseCase")

	oCatList.removeChild oCatList.selectSingleNode("APCat")
	nxml.save(Server.MapPath("/HTSDSystem/UseCase/00APlist.xml"))



sub fillUseCase

	set nUseCase = oUseCase.cloneNode(true)
	nUseCase.selectSingleNode("APCode").text = RS("APcode")
	nUseCase.selectSingleNode("APName").text = RS("APnameC")
	oUseCaseList.appendChild nUseCase

	set fxml = server.createObject("microsoft.XMLDOM")
	fxml.async = false
	fxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/useCase/" & RS("APcode") & ".xml")
	xv = fxml.load(LoadXML)
	if fxml.parseError.reason <> "" then
		Response.Write(RS("APcode") & ".xml parseError on line " &  fxml.parseError.line)
		Response.Write("<BR>Reason: " &  fxml.parseError.reason & "<HR>")
    	exit sub
	end if
	
	for each xv in fxml.selectNodes("UseCase/Version[HighLevelSpec]")
		set nVersion = nUseCase.selectSingleNode("//HighLevelSpec").cloneNode(true)
		nVersion.selectSingleNode("Date").text = nullText(xv.selectSingleNode("Date"))
		nVersion.selectSingleNode("Author").text = nullText(xv.selectSingleNode("Author"))
		if nullText(xv.selectSingleNode("//Description")) <> "" then _
			nVersion.selectSingleNode("Status").text = "Y"
		nUseCase.appendChild nVersion
	next

	for each xv in fxml.selectNodes("UseCase/Version[ExpandedSpec]")
		set nVersion = nUseCase.selectSingleNode("//ExpandedSpec").cloneNode(true)
		nVersion.selectSingleNode("Date").text = nullText(xv.selectSingleNode("Date"))
		nVersion.selectSingleNode("Author").text = nullText(xv.selectSingleNode("Author"))
		if nullText(xv.selectSingleNode("//TypicalCourseOfEvents")) <> "" then _
			nVersion.selectSingleNode("Status").text = "Y"
		nVersion.appendChild xv.selectSingleNode("ExpandedSpec/TypicalCourseOfEvents")
'		nUseCase.insertBefore nVersion, nUseCase.selectSingleNode("//ExpandedSpec")
		nUseCase.appendChild nVersion
	next

	nUseCase.removeChild nUseCase.selectSingleNode("HighLevelSpec")
	nUseCase.removeChild nUseCase.selectSingleNode("ExpandedSpec")

end sub

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
%>
