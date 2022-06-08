	<MenuBar myTreeNode="<%=myTreeNode%>">
<%
	response.write "<xRootID>"&xRootID&"</xRootID>"
	response.write "<upParent>"&upParent&"</upParent>"
	
	if cStr(xRootID)<>nullText(refModel.selectSingleNode("MenuTree")) _
			AND cStr(xRootID)<>nullText(refModel.selectSingleNode("MenuTree2")) then
		xRootID = nullText(refModel.selectSingleNode("MenuTree"))
		myTreeNode = 0
		response.write "<xRootID>"&xRootID&"</xRootID>"
		response.write "<upParent>"&upParent&"</upParent>"
	end if

  if myTreeNode=0 then
		sqlCom = "SELECT * FROM CatTreeRoot WHERE CtRootID=" & xRootID
		set RS = conn.execute(sqlCom)
%>
			<MenuCat>
				<Caption><%=deAmp(RS("PurPose"))%></Caption>
<%
		showChildren
		response.write "</MenuCat>"
  else
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.DataParent=" & upParent _
		& " AND a.CtRootID = " & PkStr(xRootID,"") _
		& " Order by a.CatShowOrder"

		set RS = conn.execute(SqlCom)
		while not RS.eof
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
			<MenuCat newWindow="<%=RS("newWindow")%>">
				<Caption><%=deAmp(RS("CatName"))%></Caption>
				<redirectURL><%=deAmp(xUrl)%></redirectURL>
<%
			if RS("CtNodeID") = cint(myTreeNode)  then
				showChildren
			end if
		response.write "</MenuCat>"
		RS.moveNext
	wend
  end if

%>
	</MenuBar>
<%
	xRootID=nullText(refModel.selectSingleNode("MenuTree2"))
	if xRootID<>"" then

	  response.write "<MenuBar2>"	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.DataLevel=1 AND a.CtRootID = " & PkStr(xRootID,"") _
		& " Order by a.CatShowOrder"

	set RS = conn.execute(SqlCom)
	while not RS.eof
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
			<MenuCat newWindow="<%=RS("newWindow")%>">
				<Caption><%=RS("CatName")%></Caption>
				<redirectURL><%=xUrl%></redirectURL>
			</MenuCat>
<%		
		RS.moveNext
	wend
%>
	</MenuBar2>
<%	end if %>
<%
sub showChildren
				  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
					& " FROM CatTreeNode AS a " _
					& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
					& " WHERE a.DataParent=" & myTreeNode & " AND a.CtRootID = " & PkStr(xRootID,"") _
					& " Order by a.CatShowOrder"
			
				set RSx = conn.execute(SqlCom)
				while not RSx.eof
						if RSx("CtUnitKind") = "K" then
					   		xURL = "kmDoit.asp?"&RSx("redirectURL")
					   	elseif RSx("CtNodeKind") = "C" then		'-- Folder
							xUrl = "np.asp?ctNode="&RSx("CtNodeID") 
						else
							if RSx("redirectURL")<> "" then
								xUrl = RSx("redirectURL")
							elseif RSx("CtUnitKind") ="2" then
								xUrl = "lp.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD")
							elseif isNumeric(RSx("iBaseDSD")) then
								xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD")
							end if
				   	   end if
%>
				<MenuItem newWindow="<%=RSx("newWindow")%>">
					<Caption><%=deAmp(RSx("CatName"))%></Caption>
					<redirectURL><%=deAmp(xURL)%></redirectURL>
				</MenuItem>
<%						RSx.moveNext
				wend
end sub
%>

