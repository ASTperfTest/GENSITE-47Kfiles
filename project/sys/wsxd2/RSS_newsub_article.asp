<% 
Response.Expires = 0
response.ContentType = "text/xml" %><?xml version="1.0"  encoding="utf-8" ?>
<rss version="2.0">
	<channel>
<%
	Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	SQL = " SELECT TOP 3 CuDTGeneric.iCUItem, CuDTGeneric.sTitle, CatTreeNode.CatName, CatTreeRoot.CtRootName, CuDTGeneric.showType, "
	SQL = SQL &	" CatTreeNode.CtRootID, CatTreeNode.CtNodeID, CuDTGeneric.xPostDate, CuDTGeneric.xURL, CuDTGeneric.fileDownLoad "
	SQL = SQL &	" FROM CuDTGeneric INNER JOIN CtUnit ON CuDTGeneric.iCTUnit = CtUnit.CtUnitID INNER JOIN "
  SQL = SQL &	" CatTreeRoot INNER JOIN CatTreeNode ON CatTreeRoot.CtRootID = CatTreeNode.CtRootID ON CtUnit.CtUnitID = CatTreeNode.CtUnitID " 
  SQL = SQL &	" WHERE (CuDTGeneric.fCTUPublic = 'Y') AND (CuDTGeneric.siteId = '2') AND (CatTreeNode.inUse = 'Y') AND (CatTreeRoot.vGroup = 'XX') AND (CatTreeRoot.inUse = 'Y') " 
  SQL = SQL &	" ORDER BY CuDTGeneric.xPostDate DESC,ximportant DESC,iCuItem DESC"

  'response.write "<aa>" & sql & "</aa>"
	Set RS = conn.execute(SQL)
	
	if not RS.eof then
		while not RS.eof 
			response.write "<item>" & VBCRLF
			response.write "<title><![CDATA[" & RS("sTitle") & "]]></title>" & VBCRLF
			response.write "<description><![CDATA[" & "[" & RS("CatName") & "]" & RS("CtRootName") & "]]></description>" & VBCRLF
			response.write "<pubDate><![CDATA[" & RS("xPostDate") & "]]></pubDate>" & VBCRLF
			if RS("showType") = "1" then
				response.write "<link><![CDATA[" & "ct.asp?xItem=" & RS("iCUItem") & "&amp;ctNode=" & RS("CtNodeID") & "&amp;mp=" & RS("CtRootID") & "]]></link>" & VBCRLF							
			elseif RS("showType") = "2" then
				response.write "<link><![CDATA[" & RS("xURL") & "]]></link>" & VBCRLF							
			elseif RS("showType") = "3" then
				response.write "<link><![CDATA[" & "/public/Data/" & RS("fileDownLoad") & "]]></link>" & VBCRLF							
			else
				response.write "<link><![CDATA[]]></link>" & VBCRLF			
			end if
			
			response.write "</item>" & VBCRLF
 		RS.movenext
		wend
	end if	
''	Conn.Close
%>
	</channel>
</rss>
