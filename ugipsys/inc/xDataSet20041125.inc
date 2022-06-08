<%
sub processXDataSet
 	xTreeNode = nullText(xDataSet.selectSingleNode("DataNode"))
 	sql = "SELECT t.*, b.sBaseTableName " _
 		& " FROM CatTreeNode AS t LEFT JOIN CtUnit AS u ON u.CtUnitID=t.CtUnitID " _
 		& " LEFT JOIN BaseDSD AS b ON b.iBaseDSD=u.iBaseDSD " _
 		& " WHERE CtNodeID IN (" & xTreeNode & ")"
 	set RS = conn.execute(sql)
  if not RS.eof then
 	xDataNode = RS("CtUnitID")
 	xBaseTableName = lCase(RS("sBaseTableName"))
 	response.write "<" & nullText(xDataSet.selectSingleNode("DataLable")) _
  		& " xNode=""" & RS("CtNodeID") & """ xUnit=""" & xDataNode & """>"
  	response.write "<Caption>" & nullText(xDataSet.selectSingleNode("DataRemark")) & "</Caption>"
  		
	HeaderCount = nullText(xDataSet.selectSingleNode("SqlTop"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount
	
	sql = "SELECT " & HeaderCount & " htx.*, ghtx.*, xr1.deptName, u.CtUnitName " _
		& " FROM CuDTGeneric AS htx JOIN " & xBaseTableName & " AS ghtx ON htx.iCuItem=ghtx.giCuItem" _
		& " LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept" _
		& " LEFT JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
		& " WHERE htx.fCTUPublic='Y' " _ 
		& " AND (htx.avEnd is NULL OR htx.avEnd >=" & pkStr(date(),")") 
'		& " LEFT JOIN CuDTwImg AS img ON img.giCuItem=htx.iCuItem" _
	if xDataNode<>"" then		sql = sql & " AND iCtUnit=" & xDataNode
	SqlCondition = nullText(xDataSet.selectSingleNode("SqlCondition"))
	if SqlCondition<>"" then	sql = sql & " AND " & SqlCondition
	SqlOrderBy = nullText(xDataSet.selectSingleNode("SqlOrderBy"))
	if SqlOrderBy<>"" then	sql = sql & " ORDER BY " & SqlOrderBy
	response.write "<sql>" & sql & "</sql>"
	set RS = conn.execute(sql)
	while not RS.eof
    	xURL = "ct.asp?xItem=" & RS("iCuItem") & "&amp;ctNode=" & xTreeNode
    	if RS("ibaseDSD") = 2 then	xURL = deAmp(RS("xURL"))
    	if xBaseTableName = "adrotate" then	xURL = deAmp(RS("xURL"))
    	if RS("showType") = 2 then	xURL = deAmp(RS("xURL"))
    	if RS("showType") = 3 then	xURL = "public/Data/" & RS("fileDownLoad")
%>
		<Article iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>">
			<Caption><%=RS("sTitle")%></Caption>
			<Content><![CDATA[<%=left(RS("xBody"),120)%>]]></Content>
			<PostDate><%=RS("xPostDate")%></PostDate>
 			<DeptName><%=RS("deptName")%></DeptName>
			<TopCat><%=RS("TopCat")%></TopCat>
			<CtUnitName><%=deAmp(RS("CtUnitName"))%></CtUnitName>
			<xURL><%=deAmp(xURL)%></xURL>
<%		if not isNull(RS("xImgFile")) then %>
			<xImgFile>public/Data/<%=RS("xImgFile")%></xImgFile>
<%		end if %>
		</Article>
<%		
		RS.moveNext
	wend

	set oxml = server.createObject("microsoft.XMLDOM")
	oxml.async = false
	oxml.setProperty("ServerHTTPRequest") = true	

	for each xRss in xDataSet.selectNodes("RSS")
	  xv = oxml.load(nullText(xRSS.selectSingleNode("URL")))
	  if xv = true then
	  	RssTop = cint(nullText(xRSS.selectSingleNode("RssTop")))
	  	xRssCount = 0
	  	for each xItem in oxml.selectNodes("rss/channel/item")
	  	  if xRssCount < RssTop then
		  	  xRssCount = xRssCount + 1
%>
		<Article iCuItem="0" newWindow="Y">
			<Caption><%=nullText(xItem.selectSingleNode("title"))%></Caption>
			<Content><![CDATA[<%=nullText(xItem.selectSingleNode("description"))%>]]></Content>
			<PostDate><%=nullText(xItem.selectSingleNode("pubDate"))%></PostDate>
			<xURL><%=nullText(xItem.selectSingleNode("link"))%></xURL>
		</Article>
<%		
		  end if
	  	next
	  end if
	next
  	response.write "</" & nullText(xDataSet.selectSingleNode("DataLable")) & ">"
  end if
end sub
%>
