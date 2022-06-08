<%
sub processXDataSet
 	xTreeNode = nullText(xDataSet.selectSingleNode("DataNode"))
	if xTreeNode = "" then	xTreeNode = 0
 	sql = "SELECT t.*, b.sBaseTableName " _
 		& " FROM CatTreeNode AS t LEFT JOIN CtUnit AS u ON u.CtUnitID=t.CtUnitID " _
 		& " LEFT JOIN BaseDSD AS b ON b.iBaseDSD=u.iBaseDSD " _
 		& " WHERE CtNodeID IN (" & xTreeNode & ")"
 	if request("debug")<>"" then response.write "<xNodesql><![CDATA[" & sql & "]]></xNodesql>"
 	set RS = conn.execute(sql)
  if not RS.eof then
 	xDataNode = RS("CtUnitID")
 	xBaseTableName = lCase(RS("sBaseTableName"))
 	dCondition = trim(RS("dCondition") & " ")
 	response.write "<" & nullText(xDataSet.selectSingleNode("DataLable")) _
  		& " xNode=""" & RS("CtNodeID") & """ xUnit=""" & xDataNode & """>"	
  	response.write "<ContentStyle>" & nullText(xDataSet.selectSingleNode("ContentStyle")) & "</ContentStyle>"
   	response.write "<DataSrc>" & nullText(xDataSet.selectSingleNode("DataSrc")) & "</DataSrc>"
   	response.write "<PicWidth>" & nullText(xDataSet.selectSingleNode("PicWidth")) & "</PicWidth>"
   	response.write "<PicHeight>" & nullText(xDataSet.selectSingleNode("PicHeight")) & "</PicHeight>"
	'主題館新增節點 start
	IsTitle = nullText(xDataSet.selectSingleNode("IsTitle"))
	if IsTitle ="" then IsTitle="Y"
	IsPic = nullText(xDataSet.selectSingleNode("IsPic"))
	if IsPic = "" then IsPic="Y"
	IsPostDate = nullText(xDataSet.selectSingleNode("IsPostDate"))
	if IsPostDate ="" then IsPostDate="Y"
	IsExcerpt = nullText(xDataSet.selectSingleNode("IsExcerpt"))
	if IsExcerpt="" then IsExcerpt="Y"
	ShowStyle = nullText(xDataSet.selectSingleNode("ShowStyle"))
	if ShowStyle="" then ShowStyle="Style1"
	
	response.write "<IsTitle>" & IsTitle & "</IsTitle>"
	response.write "<IsPic>" & IsPic & "</IsPic>"
	response.write "<IsPostDate>" & IsPostDate & "</IsPostDate>"
	response.write "<IsExcerpt>" & IsExcerpt & "</IsExcerpt>"
	response.write "<ShowStyle>" & ShowStyle & "</ShowStyle>"	
	'主題館新增節點 end
 	ContentStyle = nullText(xDataSet.selectSingleNode("ContentStyle"))	
  	response.write "<Caption><![CDATA[" & nullText(xDataSet.selectSingleNode("DataRemark")) & "]]></Caption>"
  		
	MprefDataBlock = nullText(xDataSet.selectSingleNode("refDataBlock"))
'	MprefHeaderrandom = nullText(xDataSet.selectSingleNode("refIsrandom"))
'	MprefSqlCondition = nullText(xDataSet.selectSingleNode("refSqlCondition"))
'	MprefHeaderCount = nullText(xDataSet.selectSingleNode("refHeaderCount"))
'	MprefSqlOrderBy = nullText(xDataSet.selectSingleNode("refSqlOrderBy"))

	ContentLength = nullText(xDataSet.selectSingleNode("ContentLength"))
	if ContentLength="" then	ContentLength=120
	HeaderCount = nullText(xDataSet.selectSingleNode("SqlTop"))
	Headerrandom = nullText(xDataSet.selectSingleNode("Israndom"))
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount
	
	sql = "SELECT " & HeaderCount & " htx.*, xr1.deptName, xr1.edeptName, u.CtUnitName, n.CtNodeID " 
	if Headerrandom="Y" then sql=sql & ", RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3771) " _
			& " * RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3) as ra "
'	if xBaseTableName<>"" then	sql = sql & "JOIN " & xBaseTableName & " AS ghtx ON htx.iCuItem=ghtx.giCuItem"
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

	'start 2009.07.30 by ivy
	if xTreeNode=1572 then
		if request("memID")="" then 
			'預設日出日落顯示台北區域
			sql = sql & " AND xKeyword in('台北','') "
		else
			'使用者登入後，顯示該使用者的所在區域			
			sql = sql & " AND xKeyword =CASE (SELECT keyword FROM member WHERE account='" & request("memID") & "') WHEN '' THEN '台北' ELSE (SELECT keyword FROM member WHERE account='" & request("memID") & "') END"
		end if
	end if	
	'end 2009.07.30 by ivy

	if dCondition <> "" and not isnull(dCondition) then _
		sql = sql & " AND " & replace(dCondition, "CuDTGeneric", "htx")

	SqlOrderBy = nullText(xDataSet.selectSingleNode("SqlOrderBy"))
	'有加order by條件就用原本的 沒加就用xPostDate
	if SqlOrderBy<>"" then	
		sql = sql & " ORDER BY " & SqlOrderBy
	else
		sql = sql & " ORDER BY xPostDate DESC"
	end if
'	xtn = split(xTreeNode, ",")
'	xTreeNode = xtn(0)
	if request("debug")<>"" then 
		response.write "<xDataSetsql><![CDATA[" & sql & "]]></xDataSetsql>"
		'response.end
	end if
	'on error resume next
	set RS = conn.execute(sql)
	'if error.num > 0 then
	'	response.write "<xxxError><![CDATA[" & sql & "]]></xxxError>"
	'	error.num = 0
	'	error.clear
	'end if
	while not RS.eof
    	xURL = "ct.asp?xItem=" & RS("iCuItem") & "&amp;ctNode=" & RS("CtNodeID") & "&amp;mp=" & request("mp")
    	    	
    	if RS("ibaseDSD") = 2 then	xURL = deAmp(RS("xURL"))
    	if xBaseTableName = "adrotate" then	xURL = deAmp(RS("xURL"))
    	if RS("showType") = 2 then	xURL = deAmp(RS("xURL"))
    	if RS("showType") = 3 then	xURL = "public/Data/" & RS("fileDownLoad")
    	if RS("showType") = 5 then	xURL = "content.asp?cuItem=" & RS("refID")
	
		xBody = deHTML(RS("xBody"))
		xBody = left(xBody, ContentLength)
		if len(xBody) > ContentLength then	xBody = left(xBody, ContentLength)
%>
		<Article iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>">	
			<Caption><![CDATA[<%=deAmp(RS("sTitle")) %>]]></Caption>
			<Content><![CDATA[<%=xBody%>]]></Content>
			<ContentAll><![CDATA[<%=RS("xBody")%>]]></ContentAll>
			<ContentText><![CDATA[<%=left(deHTML(RS("xBody")),250)%>]]></ContentText>
			<Abstract><![CDATA[<%=RS("xabstract")%>]]></Abstract>
			<PostDate><%=xMonStdDay(RS("xPostDate"))%></PostDate>
 			<DeptName><![CDATA[<%=RS("deptName")%>]]></DeptName>
 			<EDeptName><![CDATA[<%=RS("edeptName")%>]]></EDeptName> 			
 			<Group><%=RS("vGroup")%></Group>
			<TopCat><%=RS("TopCat")%></TopCat>
			<CtUnitName><![CDATA[<%=deAmp(RS("CtUnitName"))%>]]></CtUnitName>
			<xURL><%=deAmp(xURL)%></xURL>
			<ShowType><%=RS("showType")%></ShowType>
			<FileDownLoad><![CDATA[<%=deAmp(RS("fileDownLoad"))%>]]></FileDownLoad>
<%		if not isNull(RS("xImgFile")) then %>
			<xImgFile><![CDATA[public/Data/<%=RS("xImgFile")%>]]></xImgFile>
<%		end if %>

<%	if MprefDataBlock <> "" then 
		xsql = "SELECT Top 1 htx.* " 
		xsql=xsql & ", RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3771) " _
			& " * RAND (((icuitem+DATEPART(ms, GETDATE()))*100)%3) as ra "
		xsql=xsql &   " FROM CuDTGeneric AS htx" _
			& " WHERE htx.fCTUPublic='Y' AND ictunit=" & MprefDataBlock _
			& " AND (htx.avEnd is NULL OR htx.avEnd >=" & pkStr(date(),")") _
			& " AND (htx.avBegin is NULL OR htx.avBegin <=" & pkStr(date(),")") 
		xsql = xsql & " ORDER BY ra"
		
		set xRS = conn.execute(xsql)
		while not xRS.eof

%>	
		<refDataBlock iCuItem="<%=xRS("iCuItem")%>" newWindow="<%=xRS("xNewWindow")%>">
			<Caption><![CDATA[<%=deAmp(xRS("sTitle"))%>]]></Caption>
<%		if not isNull(xRS("xImgFile")) then %>
			<xImgFile><![CDATA[public/Data/<%=xRS("xImgFile")%>]]></xImgFile>
<%		end if %>
		</refDataBlock>
<%
		xRS.moveNext
	wend
	end if
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
			<Caption><![CDATA[<%=nullText(xItem.selectSingleNode("title"))%>]]></Caption>
			<CaptionEn><![CDATA[<%=nullText(xItem.selectSingleNode("titleEn"))%>]]></CaptionEn>
			<Content><![CDATA[<%=deHTML(nullText(xItem.selectSingleNode("description")))%>]]></Content>
			<PostDate><%=nullText(xItem.selectSingleNode("pubDate"))%></PostDate>
			<xURL><%=deAmp(nullText(xItem.selectSingleNode("link")))%></xURL>
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