	<MenuBar myTreeNode="<%=myTreeNode%>">
<%
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	
	  SqlCom = "SELECT b.CatName AS xCat, b.ctNodeID AS xNodeID, a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a JOIN CatTreeNode AS b ON b.CtNodeID=a.DataParent" _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.DataLevel=2 AND a.CtRootID = " & PkStr(xRootID,"") _
		& " Order by b.CatShowOrder, a.CatShowOrder"

	response.write"<sql>"& SqlCom &"</sql>"
	set RS = conn.execute(SqlCom)
	tabindex = 0
	divindex = 0
	styletop = 235
	xCat = ""
	while not RS.eof
	        tabindex = tabindex + 1
	        divindex = divindex + 1	    
	        if RS("CtUnitKind") = "K" then
			xURL = "kmDoit.asp?"&RS("redirectURL")
		elseif RS("CtNodeKind") = "C" then		'-- Folder
			xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;mp=" & request("mp")
		else
			if RS("redirectURL")<> "" then
				xUrl = RS("redirectURL")
			elseif RS("CtUnitKind") ="2" then
				xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
			elseif isNumeric(RS("iBaseDSD")) then
				xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
			end if
	   	end if
		if RS("xCat") <> xCat then 
			if xCat<>"" then
				styletop = styletop + 28
%>	
		</MenuCat> 
<%			end if
			xCat = RS("xCat")			
%>
		<MenuCat>
			<Caption><%=xCat%></Caption>
			<CtNode><%=deAmp(RS("xNodeID"))%></CtNode>
			<redirectURL><%=deAmp(xUrl)%></redirectURL>
			<tabindex><%=tabindex%></tabindex>
			<divindex>div<%=divindex%></divindex>
			<Mover>MM_showHideLayers('div<%=divindex%>','','show')</Mover>
			<Mout>MM_showHideLayers('div<%=divindex%>','','hide')</Mout>
			<styletop><%=styletop%></styletop>
<%		end if 	              			
%>
				<MenuItem newWindow="<%=RS("newWindow")%>">
					<Caption><%=deAmp(RS("CatName"))%></Caption>
					<redirectURL><%=deAmp(xURL)%></redirectURL>
<%
		 if RS("DataParent") = cint(myTreeNode)  then
		 	styletop = styletop + 21          	
	         end if   
%>		
				</MenuItem>
<%		
		RS.moveNext
	wend
	if xCat<>"" then	response.write "</MenuCat>"
%>
	</MenuBar>
<%
	xRootID2=nullText(refModel.selectSingleNode("MenuTree2"))
	if xRootID2<>"" then

	  response.write "<MenuBar2>"	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.DataLevel=1 AND a.CtRootID = " & PkStr(xRootID2,"") _
		& " Order by a.CatShowOrder"

	set RS = conn.execute(SqlCom)
	while not RS.eof
			if RS("CtUnitKind") = "K" then
		   		xURL = "kmDoit.asp?"&RS("redirectURL")
		   	elseif RS("CtNodeKind") = "C" then		'-- Folder
				xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;mp=" & request("mp")
			else
				if RS("redirectURL")<> "" then
					xUrl = RS("redirectURL")
				elseif RS("CtUnitKind") ="2" then
					xUrl = "lp.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
				elseif isNumeric(RS("iBaseDSD")) then
					xUrl = "np.asp?ctNode="&RS("CtNodeID") & "&amp;CtUnit=" & RS("CtUnitID") & "&amp;BaseDSD=" & RS("iBaseDSD") & "&amp;mp=" & request("mp")
				end if
	   	   end if
%>
			<MenuCat newWindow="<%=RS("newWindow")%>">
				<Caption><%=deAmp(RS("CatName"))%></Caption>
				<redirectURL><%=xUrl%></redirectURL>
			</MenuCat>
<%		
		RS.moveNext
	wend
%>
	</MenuBar2>
<%	end if %>
