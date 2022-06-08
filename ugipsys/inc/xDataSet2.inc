<%
sub processXDataSet
 	xTreeNode = nullText(xDataSet.selectSingleNode("DataNode"))
 	sql = "SELECT t.*, b.sBaseTableName " _
 		& " FROM CatTreeNode AS t LEFT JOIN CtUnit AS u ON u.CtUnitID=t.CtUnitID " _
 		& " LEFT JOIN BaseDSD AS b ON b.iBaseDSD=u.iBaseDSD " _
 		& " WHERE CtNodeID IN (" & xTreeNode & ")"
 	if request("debug")<>"" then response.write "<xNodesql><![CDATA[" & sql & "]]></xNodesql>"
 	set RS = conn.execute(sql)
  if not RS.eof then
 	xDataNode = RS("CtUnitID")
 	xBaseTableName = lCase(RS("sBaseTableName"))
 	response.write "<" & nullText(xDataSet.selectSingleNode("DataLable")) _
  		& " xNode=""" & RS("CtNodeID") & """ xUnit=""" & xDataNode & """>"
  	response.write "<Caption>" & nullText(xDataSet.selectSingleNode("DataRemark")) & "</Caption>"
  		
	HeaderCount = nullText(xDataSet.selectSingleNode("SqlTop"))
	Headerrandom = nullText(xDataSet.selectSingleNode("Israndom"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount
	
	sql = "SELECT " & HeaderCount & " htx.*, xr1.deptName, u.CtUnitName, n.CtNodeID " 
	if Headerrandom<>"" then sql=sql & ", RAND (((icuitem+DATEPART(ms, GETDATE()))*100000)%3771) as ra" 
		sql=sql &   " FROM CuDTGeneric AS htx " _
			& " JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
			& " JOIN CatTreeNode AS n ON n.CtUnitID=htx.iCtUnit" _
			& " LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept" _
			& " WHERE htx.fCTUPublic='Y' " _
			& " AND (htx.avEnd is NULL OR htx.avEnd >=" & pkStr(date(),")") _
			& " AND (htx.avBegin is NULL OR htx.avBegin <=" & pkStr(date(),")") 
'		& " LEFT JOIN CuDTwImg AS img ON img.gicuitem=htx.iCuItem" _
'	if xDataNode<>"" then		sql = sql & " AND iCtUnit=" & xDataNode
	if xTreeNode<>"" then		sql = sql & " AND n.CtNodeID IN (" & xTreeNode & ")"
		
	SqlCondition = nullText(xDataSet.selectSingleNode("SqlCondition"))
	if SqlCondition<>"" then	sql = sql & " AND " & SqlCondition
	SqlOrderBy = nullText(xDataSet.selectSingleNode("SqlOrderBy"))
	if SqlOrderBy<>"" then	sql = sql & " ORDER BY " & SqlOrderBy
'	response.write "<SQL><![CDATA["&sql&"]]></SQL>"
'	xtn = split(xTreeNode, ",")
'	xTreeNode = xtn(0)
	if request("debug")<>"" then response.write "<xDataSetsql><![CDATA[" & sql & "]]></xDataSetsql>"
	set RS = conn.execute(sql)
	while not RS.eof
    	xURL = "ct.asp?xItem=" & RS("iCuItem") & "&amp;ctNode=" & RS("CtNodeID")
    	if RS("ibaseDSD") = 2 then	xURL = deAmp(RS("xURL"))
    	if xBaseTableName = "adrotate" then	xURL = deAmp(RS("xURL"))
    	if RS("showType") = 2 then	xURL = deAmp(RS("xURL"))
    	if RS("showType") = 3 then	xURL = "public/Data/" & RS("fileDownLoad")
    	if RS("showType") = 5 then	xURL = "content.asp?cuItem=" & RS("refID")
%>
		<Article iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>">
			<Caption><%=deAmp(RS("sTitle"))%></Caption>
			<Content><![CDATA[<%=left(deHTML(RS("xBody")),120)%>]]></Content>
			<ContentAll><![CDATA[<%=RS("xBody")%>]]></ContentAll>
			<PostDate><%=RS("xPostDate")%></PostDate>
 			<DeptName><%=RS("deptName")%></DeptName>
 			<Group><%=RS("vGroup")%></Group>
			<TopCat><%=RS("TopCat")%></TopCat>
			<CtUnitName><%=deAmp(RS("CtUnitName"))%></CtUnitName>
			<xURL><%=deAmp(xURL)%></xURL>
<%		if not isNull(RS("xImgFile")) then %>
			<xImgFile>public/Data/<%=RS("xImgFile")%></xImgFile>
<%		end if %>
<%
	xCtNodeID=nullText(xDataSet.selectSingleNode("CtNodeID"))
	response.write "<CtNodeID>" & xCtNodeID & "</CtNodeID>"

%>
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
			<Content><![CDATA[<%=deHTML(nullText(xItem.selectSingleNode("description")))%>]]></Content>
			<PostDate><%=nullText(xItem.selectSingleNode("pubDate"))%></PostDate>
			<xImgFile><%=nullText(xItem.selectSingleNode("image"))%></xImgFile>
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

function deHTML(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deHTML=""
  	exit function
  end if
	xs = ReplaceTest("<[^>]*>", "", xs)	'-- <.......>
  	xs = replace(xs,vbCRLF&vbCRLF,"<P>")
  	xs = replace(xs,vbCRLF,"<BR>")
  	deHTML = replace(xs,chr(10),"<BR>")
end function

Function ReplaceTest(patrn, replStr, str1)
  Dim regEx               ' Create variables.
'  str1 = "The quick brown fox jumped over the lazy dog."
  Set regEx = New RegExp            ' Create regular expression.
  regEx.Pattern = patrn            ' Set pattern.
  regEx.IgnoreCase = True            ' Make case insensitive.
  regEx.Global = True   ' Set global applicability.
  ReplaceTest = regEx.Replace(str1, replStr)   ' Make replacement.
End Function


%>
