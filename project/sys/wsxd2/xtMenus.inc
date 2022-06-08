	<MenuBar>
<%
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	
	  SqlCom = "SELECT b.CtRootName AS xCat, a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a JOIN CatTreeRoot AS b ON b.CtRootID=a.CtRootID" _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND b.inUse='Y'" _
		& " Order by b.CtRootID, a.CatShowOrder"

	set RS = conn.execute(SqlCom)
	xCat = ""
	while not RS.eof
		if RS("xCat") <> xCat then 
			if xCat<>"" then
%>	
		</MenuCat> 
<%			end if
			xCat = RS("xCat")
%>
		<MenuCat>
			<Caption><%=xCat%></Caption>
<%		end if 
						if RS("CtUnitKind") = "K" then
					   		xURL = "kmDoit.asp?"&RS("redirectURL")
					   	elseif RS("CtNodeKind") = "C" then		'-- Folder
							xUrl = "np.asp?ctNode="&RS("CtNodeID") 
						else
							if RS("redirectURL")<> "" then
								xUrl = RS("redirectURL")
							elseif RS("CtUnitKind") ="2" then
								xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")
							elseif isNumeric(RS("iBaseDSD")) then
								xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD")
							end if
				   	   end if
%>
				<MenuItem newWindow="<%=RS("newWindow")%>">
					<Caption><%=deAmp(RS("CatName"))%></Caption>
					<redirectURL><%=deAmp(xURL)%></redirectURL>
				</MenuItem>
<%		
		RS.moveNext
	wend
	if xCat<>"" then	response.write "</MenuCat>"
%>
	</MenuBar>
