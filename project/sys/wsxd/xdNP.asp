<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#Include virtual = "/inc/xDataSet.inc" -->
<% 
function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

npURL = ""

sql = "SELECT n.CtUnitID, n.CtNodeKind, n.CtNodeNPKind, u.iBaseDSD, u.CtUnitKind" _
	& " FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON n.CtUnitID=u.CtUnitID" _
	& " WHERE n.CtNodeID=" & pkStr(request("ctNode"),"")
set RS = conn.execute(sql)
if RS("CtNodeKind") = "U" then
	if RS("CtUnitKind") = "1" then
		sql = "SELECT TOP 1 * FROM CuDTGeneric WHERE iCtUnit=" & RS("CtUnitID") & " and fctupublic = 'Y'" _
			& " ORDER BY xImportant DESC"
		set RSx = conn.execute(sql)
		if not RSx.eof then
			npURL = "ct.asp?xItem=" & RSx("iCuItem") & "&CtNode=" & request("ctNode")& "&mp=" & request("mp")
		else
			npURL = "ct.asp?xItem=&CtNode=" & request("ctNode")& "&mp=" & request("mp")
		end if
	else
			npURL = "lp.asp?CtNode=" & request("CtNode") & "&CtUnit=" & RS("CtUnitID") & "&BaseDSD=" & RS("iBaseDSD")& "&mp=" & request("mp")
	end if
	response.write "<npURL>"&deAmp(npURL)&"</npURL>"
	'response.write "<sql>"&sql&"</sql>"
elseif IsNull(RS("CtNodeNPKind")) then
	sql = "SELECT * FROM CatTreeNode WHERE DataParent=" & pkStr(request("ctNode"),"") _
		& " ORDER BY CatShowOrder"
	set RS = conn.execute(sql)
	npURL = "np.asp?ctNode=" & RS("CtNodeID")& "&mp=" & request("mp")
	response.write "<npURL>"&deAmp(npURL)&"</npURL>"
else

	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	htPageDom.setProperty("ServerHTTPRequest") = true	

	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\xdmp" & request("mp") & ".xml"
	xv = htPageDom.load(LoadXML)
	if htPageDom.parseError.reason <> "" then 
		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
		Response.End()
	end if
  	set refModel = htPageDom.selectSingleNode("//MpDataSet")
	myTreeNode = request("ctNode")
%>
	<xslpage>np</xslpage>
	<%

	sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(request("ctNode"),"")
	set RS = conn.execute(sql)
	xRootID = RS("CtRootID")
	xCtUnitName = RS("CatName")
	xPathStr = "<xPathNode Title=""" & deAmp(xCtUnitName) & """ xNode=""" & RS("ctNodeID") & """ />"
	xParent = RS("DataParent")
'	xParent = RS("CtNodeID")
	myParent = xParent
	xLevel = RS("DataLevel") - 1
	if RS("CtNodeKind") <> "C" then
		xLevel = xLevel -1
		myTreeNode = xParent
	end if
	upParent = 0
	myupParent = RS("CtNodeID")
	
	while xParent<>0
		sql = "SELECT * FROM CatTreeNode where CtNodeID=" & pkstr(xParent,"")
		set RS = conn.execute(sql)
		if RS("DataLevel") = xLevel then	upParent = xParent
		xPathStr = "<xPathNode Title=""" & deAmp(RS("CatName")) & """ xNode=""" & RS("ctNodeID") & """ />" & xPathStr
		xParent = RS("DataParent")
	wend
	response.write "<xPath><UnitName>" & deAmp(xCtUnitName) & "</UnitName>" & xpathStr & "</xPath>"
  
	for each xDataSet in refModel.selectNodes("DataSet[ContentData='Y']")
		processXDataSet
	next
%>
	<info myTreeNode="<%=myTreeNode%>" upParent="<%=upParent%>" myParent="<%=myParent%>" myupParent="<%=myupParent%>"/>
	<MenuBar3 myTreeNode="<%=myTreeNode%>" Caption="<%=deAmp(xCtUnitName)%>">
		<%
	if isNull(xRootID) then _
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	
	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataParent=" & myupParent _
		& " AND a.CtRootID = " & PkStr(xRootID,"") _
		& " Order by a.CatShowOrder"

	set RS = conn.execute(SqlCom)
	while not RS.eof
			if RS("CtUnitKind") = "K" then
		   		xURL = "kmDoit.asp?"&RS("redirectURL")
		   	elseif RS("CtNodeKind") = "C" then		'-- Folder
				xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&mp=" & request("mp")
			else
				if RS("redirectURL")<> "" then
					xUrl = RS("redirectURL")
				elseif RS("CtUnitKind") ="2" then
					xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
				elseif isNumeric(RS("iBaseDSD")) then
					xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")& "&amp;mp=" & request("mp")
				end if
	   	   end if
%>
		<MenuCat newWindow="<%=RS("newWindow")%>">
			<Caption>
				<%=deAmp(RS("CatName"))%>
			</Caption>
			<redirectURL>
				<%=deAmp(xUrl)%>
			</redirectURL>
			<%
			if RS("CtNodeID") = cint(myTreeNode)  then
				  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
					& " FROM CatTreeNode AS a " _
					& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
					& " WHERE a.inUse='Y' AND a.DataParent=" & myTreeNode & " AND a.CtRootID = " & PkStr(xRootID,"") _
					& " Order by a.CatShowOrder"
			
				set RSx = conn.execute(SqlCom)
				while not RSx.eof
						if RSx("CtUnitKind") = "K" then
					   		xURL = "kmDoit.asp?"&RSx("redirectURL")
					   	elseif RSx("CtNodeKind") = "C" then		'-- Folder
							xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&mp=" & request("mp")
						else
							if RSx("redirectURL")<> "" then
								xUrl = RSx("redirectURL")
							elseif RSx("CtUnitKind") ="2" then
								xUrl = "lp.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD")& "&amp;mp=" & request("mp")
							elseif isNumeric(RSx("iBaseDSD")) then
								xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD")& "&amp;mp=" & request("mp")
							end if
				   	   end if
%>
			<MenuItem newWindow="<%=RSx("newWindow")%>">
				<Caption>
					<%=deAmp(RSx("CatName"))%>
				</Caption>
				<redirectURL>
					<%=deAmp(xURL)%>
				</redirectURL>
			</MenuItem>
			<%						RSx.moveNext
				wend
			end if
%>
		</MenuCat>
		<%		
		RS.moveNext
	wend
%>
	</MenuBar3>
	<!--#include file="x1Menus.inc" -->
<%end if%>
</hpMain>
