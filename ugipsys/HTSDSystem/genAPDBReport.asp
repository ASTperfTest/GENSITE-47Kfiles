<% HTProgCap="AP"
HTProgCode="HT090"
HTProgPrefix="AP" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/xdbutil.inc" -->
<%
dim xxAPcode

	set nxml = server.createObject("microsoft.XMLDOM")
	nxml.async = false
	nxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/APDBlist0.xml")
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
		if left(RS("APnameC"),2)<>"總目" then _
			fillUseCase
	
		RS.moveNext
	wend
	response.write "</TABLE>"

	if xCatID <> "" then _
		oUseCaseList.removeChild oUseCaseList.selectSingleNode("UseCase")

	oCatList.removeChild oCatList.selectSingleNode("APCat")
	nxml.save(Server.MapPath("/HTSDSystem/UseCase/00APDBlist.xml"))


sub fillUseCase

	set nUseCase = oUseCase.cloneNode(true)
	nUseCase.selectSingleNode("APCode").text = RS("APcode")
	xxAPcode = RS("APcode")
	nUseCase.selectSingleNode("APName").text = RS("APnameC")
	oUseCaseList.appendChild nUseCase

	xsql = "DELETE agrp WHERE apCode=" & pkStr(xxAPcode,"")
	conn.execute xsql


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
	
	if nullText(fxml.selectSingleNode("UseCase/Version[ExpandedSpec]/Date")) <> "" then
	  set xv = fxml.selectSingleNode("UseCase/Version[ExpandedSpec]")
		set nVersion = nUseCase.selectSingleNode("//ExpandedSpec").cloneNode(true)
		nVersion.selectSingleNode("Date").text = nullText(xv.selectSingleNode("Date"))
		nVersion.selectSingleNode("Author").text = nullText(xv.selectSingleNode("Author"))
		if nullText(xv.selectSingleNode("//TypicalCourseOfEvents")) <> "" then _
			nVersion.selectSingleNode("Status").text = "Y"

		for each tcoe in xv.selectNodes("ExpandedSpec/TypicalCourseOfEvents/Event")
'			response.write tcoe.getAttribute("ContractXML") & "z<BR> "
'			response.write tcoe.text & "z<BR> "
			fillEvent nVersion, tcoe.getAttribute("ContractXML")
		next

		nVersion.appendChild xv.selectSingleNode("ExpandedSpec/TypicalCourseOfEvents")
'		nUseCase.insertBefore nVersion, nUseCase.selectSingleNode("//ExpandedSpec")
		nUseCase.appendChild nVersion
	end if
	
	nUseCase.removeChild nUseCase.selectSingleNode("ExpandedSpec")

end sub

sub fillEvent (xVer, xFile)
on error resume next
	if xFile="" then	exit sub
	set efxml = server.createObject("microsoft.XMLDOM")
	efxml.async = false
	efxml.setProperty "ServerHTTPRequest", true
	LoadXML = server.mappath("/HTSDSystem/useCase/" & xFile)
	exv = efxml.load(LoadXML)
	if efxml.parseError.reason <> "" then
		Response.Write(RS("APcode") & ".xml parseError on line " &  efxml.parseError.line)
		Response.Write("<BR>Reason: " &  efxml.parseError.reason & "<HR>")
    	exit sub
	end if
	agrpID= mid(efxml.selectSingleNode("//ContractSpec/Name").text,5)
	if agrpID="" then	exit sub
	efxml.selectSingleNode("//ContractSpec/Name").text = agrpID
	agrpName = nulltext(efxml.selectSingleNode("//ContractSpec/ResponsibilityList/Responsibility"))
'on error GoTo  0

	xsql = "DELETE agrpFP WHERE agrpID=" & pkStr(agrpID,"")
	conn.execute xsql

	xsql = "INSERT INTO agrp(agrpID,agrpName,apCode) VALUES(" _
		& pkStr(agrpID,",") & pkstr(agrpName,",") & pkStr(xxAPcode,")")
	conn.execute xsql
	
	for each xfAccess in efxml.selectNodes("//ContractSpec/htDEntityList/htDEntity")
		xRight = 0
		FPCode = xfAccess.text
		if ucase(xfAccess.getAttribute("xRef")) = "Y" then	xRight = xRight + 1
		if ucase(xfAccess.getAttribute("xUpdate")) = "Y" then	xRight = xRight + 2
'		response.write FPcode & "==>" & xRight & "==>"
		for each pceop in efxml.selectNodes("//ContractSpec/PostConditionList/PostCondition[.='" & FPCode & "']")
'			response.write pceof.text & pceop.getAttribute("eop")
			if pceop.getAttribute("eop") = "instanceCreate" then 	xRight = xRight OR 4
			if pceop.getAttribute("eop") = "instanceDelete" then 	xRight = xRight OR 8
		next
'		response.write "=(" & xRight & ")<BR>"

		xsql = "INSERT INTO agrpFP(agrpID, FPCode, rights) VALUES(" _
			& pkStr(agrpID,",") & pkStr(FPCode,",") & xRight & ")"
		conn.execute xsql

	next
	
	xVer.appendChild efxml.selectSingleNode("//ContractSpec")


end sub


function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
%>
<HTML>
<BODY>
<SCRIPT language="vbs">
	window.navigate "APDBreport.asp"
</SCRIPT>
</BODY>
</HTML>
