
	<HitPage>
<%
	sql = " SELECT TOP 5 iCuItem, iCtNode FROM gipHitPage " & _
				" GROUP BY iCuItem, iCtNode ORDER BY COUNT(*) DESC "
	set rs = conn.execute(sql)
	
	while not rs.eof
	
		sql2 = "SELECT sTitle FROM CuDTGeneric WHERE iCuItem = " & rs("iCuItem")
		set rs2 = conn.execute(sql2)
		if not rs2.eof then	
			sTitle = rs2("sTitle")
			sTitle = Replace(sTitle, "<", "&lt;")
			sTitle = Replace((sTitle, ">", "&gt;")
%>
			<Article iCuItem="<%= rs("iCuItem") %>" iCtNode="<%= rs("iCtNode") %>">
				<Caption><%=sTitle%></Caption>
			</Article>
<%
		end if		
		rs.movenext
	wend
%>
	</HitPage>
	
	<HitPageAll>
<%
	sql = " SELECT TOP 20 iCuItem, iCtNode FROM gipHitPage " & _
				" GROUP BY iCuItem, iCtNode ORDER BY COUNT(*) DESC "
	set rs = conn.execute(sql)
	while not rs.eof
		sql2 = "SELECT sTitle FROM CuDTGeneric WHERE iCuItem = " & rs("iCuItem")
		set rs2 = conn.execute(sql2)
		if not rs2.eof then	
			sTitle = rs2("sTitle")
			sTitle = Replace(sTitle, "<", "&lt;")
			sTitle = Replace((sTitle, ">", "&gt;")
%>
			<Article iCuItem="<%= rs("iCuItem") %>" iCtNode="<%= rs("iCtNode") %>">
				<Caption><%=sTitle%></Caption>
			</Article>
<%
		end if		
		rs.movenext
	wend
%>
	</HitPageAll>