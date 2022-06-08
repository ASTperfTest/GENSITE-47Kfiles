	<MenuBar myTreeNode="<%=myTreeNode%>">
<%
	if isNull(xRootID) then _
	xRootID=nullText(refModel.selectSingleNode("MenuTree"))
	
	sqlstring = "select CatMemo_Disable from nodeinfo where ctrootid =" & PkStr(xRootID,"")
	set rsDis = conn.Execute(sqlstring)
	memo_disable = false
	if not rsDis.eof then
		if rsDis("CatMemo_Disable") <> "" or rsDis("CatMemo_Disable") <> NULL then
			memo_disable = rsDis("CatMemo_Disable")
		end if
	end if
	
	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataParent=" & upParent _
		& " AND a.CtRootID = " & PkStr(xRootID,"") _
		& " Order by a.CatShowOrder"

	set RS = conn.execute(SqlCom)
	tabindex = 0
	divindex = 0
	styletop = 207
	while not RS.eof
	                tabindex = tabindex + 1
	                divindex = divindex + 1
	                styletop = styletop + 28
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
<%
	 if not memo_disable then 
%>
				<CaptionTip><%=deAmp(RS("CatNameMemo"))%></CaptionTip>
<% else %> 		
				<CaptionTip><%=deAmp(RS("CatName"))%></CaptionTip>
<% end if %> 
				<CtNode><%=RS("CtNodeID")%></CtNode>
				<redirectURL><%=deAmp(xUrl)%></redirectURL>
				<tabindex><%=tabindex%></tabindex>
				<divindex>div<%=divindex%></divindex>
				<Mover>MM_showHideLayers('div<%=divindex%>','','show')</Mover>
				<Mout>MM_showHideLayers('div<%=divindex%>','','hide')</Mout>
				<styletop><%=styletop%></styletop>
<%
			
				  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
					& " FROM CatTreeNode AS a " _
					& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
					& " WHERE a.inUse='Y' and a.DataParent=" & RS("CtNodeID") & " AND a.CtRootID = " & PkStr(xRootID,"") _
					& " Order by a.CatShowOrder"
			
				set RSx = conn.execute(SqlCom)
				while not RSx.eof
					tabindex = tabindex + 1
					if RS("CtNodeID") = cint(myTreeNode)  then
						styletop = styletop + 26
					end if
						if RSx("CtUnitKind") = "K" then
					   		xURL = "kmDoit.asp?"&RSx("redirectURL")
					   	elseif RSx("CtNodeKind") = "C" then		'-- Folder
							xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&amp;mp=" & request("mp")
						else
							if RSx("redirectURL")<> "" then
								xUrl = RSx("redirectURL")
							elseif RSx("CtUnitKind") ="2" then
								xUrl = "lp.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD") & "&amp;mp=" & request("mp")
							elseif isNumeric(RSx("iBaseDSD")) then
								xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD") & "&amp;mp=" & request("mp")
							end if
				   	   end if
%>
				<MenuItem newWindow="<%=RSx("newWindow")%>">
					<Caption><%=deAmp(RSx("CatName"))%></Caption>
<%
	 if not memo_disable then 
%>
					<CaptionTip><%=deAmp(RSx("CatNameMemo"))%></CaptionTip>
<% else %> 		
					<CaptionTip><%=deAmp(RSx("CatName"))%></CaptionTip>
<% end if %> 
					<redirectURL><%=deAmp(xURL)%></redirectURL>
					<tabindex><%=tabindex%></tabindex>
				</MenuItem>
<%						RSx.moveNext
				wend
			
%>
			</MenuCat>
<%		
		RS.moveNext
	wend
%>
	</MenuBar>

<%
	xRootID=nullText(refModel.selectSingleNode("MenuTree1"))
	if xRootID<>"" then

	  response.write "<MenuBar1>"	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataLevel=1 AND a.CtRootID = " & PkStr(xRootID,"") _
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
				<Caption><%=RS("CatName")%></Caption>
				<redirectURL><%=xUrl%></redirectURL>
<%
			if RS("CtNodeID") = cint(myTreeNode)  then
				  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
					& " FROM CatTreeNode AS a " _
					& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
					& " WHERE a.inUse='Y' and a.DataParent=" & myTreeNode & " AND a.CtRootID = " & PkStr(xRootID,"") _
					& " Order by a.CatShowOrder"
			
				set RSx = conn.execute(SqlCom)
				while not RSx.eof
						if RSx("CtUnitKind") = "K" then
					   		xURL = "kmDoit.asp?"&RSx("redirectURL")
					   	elseif RSx("CtNodeKind") = "C" then		'-- Folder
							xUrl = "np.asp?ctNode="&RSx("CtNodeID")  & "&amp;mp=" & request("mp")
						else
							if RSx("redirectURL")<> "" then
								xUrl = RSx("redirectURL")
							elseif RSx("CtUnitKind") ="2" then
								xUrl = "lp.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD") & "&amp;mp=" & request("mp")
							elseif isNumeric(RSx("iBaseDSD")) then
								xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD") & "&amp;mp=" & request("mp")
							end if
				   	   end if
%>
				<MenuItem newWindow="<%=RSx("newWindow")%>">
					<Caption><%=deAmp(RSx("CatName"))%></Caption>
					<redirectURL><%=deAmp(xURL)%></redirectURL>
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
	</MenuBar1>
<%	end if %>

<%
	xRootID=nullText(refModel.selectSingleNode("MenuTree2"))
	if xRootID<>"" then

	  response.write "<MenuBar2>"	
	  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
		& " FROM CatTreeNode AS a " _
		& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
		& " WHERE a.inUse='Y' AND a.DataLevel=1 AND a.CtRootID = " & PkStr(xRootID,"") _
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
				<Caption><%=RS("CatName")%></Caption>
				<redirectURL><%=xUrl%></redirectURL>
<%
				  SqlCom = "SELECT a.*, c.redirectURL, c.newWindow, c.iBaseDSD, c.CtUnitKind " _
					& " FROM CatTreeNode AS a " _
					& " LEFT JOIN CtUnit AS c ON c.ctUnitID=a.CtUnitID" _
					& " WHERE a.inUse='Y' and a.DataParent=" & RS("ctNodeId") & " AND a.CtRootID = " & PkStr(xRootID,"") _
					& " Order by a.CatShowOrder"
			
				set RSx = conn.execute(SqlCom)
				while not RSx.eof
						if RSx("CtUnitKind") = "K" then
					   		xURL = "kmDoit.asp?"&RSx("redirectURL")
					   	elseif RSx("CtNodeKind") = "C" then		'-- Folder
							xUrl = "np.asp?ctNode="&RSx("CtNodeID")  & "&amp;mp=" & request("mp")
						else
							if RSx("redirectURL")<> "" then
								xUrl = RSx("redirectURL")
							elseif RSx("CtUnitKind") ="2" then
								xUrl = "lp.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD") & "&amp;mp=" & request("mp")
							elseif isNumeric(RSx("iBaseDSD")) then
								xUrl = "np.asp?ctNode="&RSx("CtNodeID") & "&amp;CtUnit=" & RSx("CtUnitID") & "&amp;BaseDSD=" & RSx("iBaseDSD") & "&amp;mp=" & request("mp")
							end if
				   	   end if
%>
				<MenuItem newWindow="<%=RSx("newWindow")%>">
					<Caption><%=deAmp(RSx("CatName"))%></Caption>
					<redirectURL><%=deAmp(xURL)%></redirectURL>
				</MenuItem>
<%						RSx.moveNext
				wend
%>
			</MenuCat>
<%		
		RS.moveNext
	wend
%>
	</MenuBar2>
<%	end if  %>
	<SubjectBeaanCursor>
	<%	
		'四季豆仔游標使用
		response.write "<CursorOpen>"
		showPic = "2"
		if request("memID") <> "" then 
			SqlCom = "select ShowCursorIcon from Member Where account = '" & request("memID") & "'"
			Set loginrs = conn.execute(SqlCom)
			If Not loginrs.Eof Then
				If Not IsNull(loginrs("ShowCursorIcon")) Then
					showPic = loginrs("ShowCursorIcon")
				End If
			End If
		End If
			
		If showPic = "2" then
			SqlCom = " select stitle from CuDTGeneric where icuitem =" & Application("ShowCursorIconId")
			set RSOP = conn.execute(SqlCom)
			if (Not IsNull(RSOP("sTitle")) ) and RSOP("sTitle") = 1 then
				showPic = "1"
			else
				showPic = "0"
			End If
		End IF
			
		if showPic = "1" Then
			response.write "1</CursorOpen>"
			SqlCom = "select sTitle from CuDTGeneric where ictunit=302 "
			set RSB = conn.execute(SqlCom)
			if Not RSB.eof then
				response.write "<CursorPic>" & RSB("sTitle") & "</CursorPic>"
			else
				response.write "<CursorPic>1</CursorPic>"
			End If
		else
			response.write "0</CursorOpen><CursorPic></CursorPic>"
		End If
	%>
	</SubjectBeaanCursor>



