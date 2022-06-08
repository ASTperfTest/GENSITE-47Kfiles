<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<%

'IEditor=session("userID") 
'IDept=session("deptID") 

Regicuitem= request("gicuitem")

orderArticle = "1"
IEditor = "hyweb"
IDept = "0"
showType = "1"
siteId = "1"
iBaseDSD = "44"
iCTUnit = "2201"

Set KMConn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'KMConn.Open session("KMODBC")
Set KMConn = Server.CreateObject("HyWebDB3.dbExecute")
KMConn.ConnectionString = session("KMODBC")
'----------HyWeb GIP DB CONNECTION PATCH----------


check			= request("check")
uncheck		=	request("uncheck")
jigcheck	=	session("jigcheck")
len1			=	Len(session("jigsql"))
aa				=	session("jigsql")

a					=	session("jigcheck")
b					=	check
c					=	uncheck

parenticuitem = request("gicuitem")



'response.write Len(c)
checkarr = split(b, ";")
uncheckarr = split(c, ";")
For i = 0 To UBound(checkarr)
   if (InStr(a,checkarr(i))>0) then

   else
	   add=checkarr(i)+";"
	   a=a+add
	end if
next


For i = 0 To UBound(uncheckarr)-1
    if (InStr(a,uncheckarr(i))>0) then
	   cut=uncheckarr(i)+";"
       a=Replace(a,cut,"")
    end if
next




session("jigcheck")=a
insert=split(session("jigcheck"), ";")
'for i = 0 to ubound(insert)
'	response.write insert(i) & "~"
'next


if request("CtRootId") = 4 then
	for i = 0 TO UBound(insert) - 1
		sql = Mid( aa, 1, len1 - 35 )
		str = insert(i)
		'-----modify by vincent on 2008-11-09-----
		'-----reason : 傳過來的值是用"-"分隔, 所以要用split來取出值, -----
		'-----最好不要用字串長度去取, 因為字串長度可能不是固定的 -----		
		set items = nothing
		items = split(str, "-")
		insert1 = items(0)
		'-----------------------------------------------------
		'-----原本的寫法-----
		'len2=Len(str)
		'bb=InStr(str,"-")
		'insert1=mid(str,1,bb-1)		
		'-----------------------------------------------------		
		if not CheckExist(insert1) then			
	    sql = sql & " AND (REPORT.REPORT_ID = '" & insert1 & "')"
      Set rs = KMConn.Execute(sql)
			' dEditDate=right(year(rs("ONLINE_DATE")),4)   &"/"&   right("0"&month(rs("ONLINE_DATE")),2)&   "/"   &   right("0"&day(rs("ONLINE_DATE")),2)   &   "   "   &   right("0"&hour(rs("ONLINE_DATE")),2)&   ":"   &   right("0"&minute(rs("ONLINE_DATE")),2)&   ":"   &   right("0"&second(rs("ONLINE_DATE")),2) 
			sql2 = "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[dEditDate],[iDept],[showType],[siteId]) "
			sql2 = sql2 & "VALUES(" & iBaseDSD & ", " & iCTUnit & ", '" & rs("SUBJECT") & "', '" & IEditor & "', GETDATE(), '" & IDept & "', "
			sql2 = sql2 & "'" & showType & "', '" & siteId & "') "
			sql2 = "set nocount on;" & sql2 & "; select @@IDENTITY as NewID"	
			Set rs2 = conn.Execute(sql2)
			gicuitem = rs2(0)
			sql6 = "INSERT INTO CuDTx7 ([giCuItem]) VALUES(" & gicuitem & ") "		   			
			conn.Execute(sql6)
			
			CtUnitId= rs("CATEGORY_ID")
			
			path = "/CatTree/CatTreeContent.aspx?ReportId={0}&DatabaseId=DB020&CategoryId={1}&ActorType=002"
			path = replace(path, "{0}", insert1)
			path = replace(path, "{1}", CtUnitId)
			
			sql2i = "INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[CtRootId],[CtUnitId],[parentIcuitem],[ArticleId],[Status],[orderArticle],[path]) "
			sql2i = sql2i & "VALUES(" & gicuitem & ", " & request("CtRootId") & ", '" & CtUnitId & "', " & Regicuitem & ", " & insert1 & ", 'Y', " & orderArticle & ", '" & path & "')"						
			conn.Execute(sql2i)
	  end if 
  next
	showDoneBox "編修成功！"
else



	for i = 0 TO UBound(insert)-1

		sql = Mid( aa, 1, len1 - 36 )		
		str = insert(i)	'-----取得id值----以"-"分隔-----			
		'-----modify by vincent on 2008-11-09-----
		'-----reason : 傳過來的值是用"-"分隔, 所以要用split來取出值, -----
		'-----最好不要用字串長度去取, 因為字串長度可能不是固定的 -----		
		items = split(str, "-")
		insert1 = items(0)
		'-----------------------------------------------------
		'-----原本的寫法-----
		'len2 = Len(str)
		'insert1 = mid(str,1,len2-4)
		'-----------------------------------------------------
		
		
		if not CheckExist(insert1) then
    	   	  
		  sql = sql & " AND (CuDTGeneric.iCUItem = " & insert1 & ") "		 

		  Set rs = conn.Execute(sql)
			'dEditDate=right(year(rs("dEditDate")),4)   &"/"&   right("0"&month(rs("dEditDate")),2)&   "/"   &   right("0"&day(rs("dEditDate")),2)   &   "   "   &   right("0"&hour(rs("dEditDate")),2)&   ":"   &   right("0"&minute(rs("dEditDate")),2)&   ":"   &   right("0"&second(rs("dEditDate")),2) 
		  sql1 = "INSERT INTO [mGIPcoanew].[dbo].[CuDTGeneric]([iBaseDSD],[iCTUnit],[sTitle],[iEditor],[dEditDate],[iDept],[showType],[siteId]) "
		  sql1 = sql1 & "VALUES(" & iBaseDSD & ", " & iCTUnit & ", '" & rs("sTitle") & "', '" & IEditor & "', GETDATE(), '" & IDept & "', "
			sql1 = sql1 & "'" & showType & "', '" & siteId & "') "
		  sql1 = "set nocount on;" & sql1 & "; select @@IDENTITY as NewID"	
		  Set rs1 = conn.Execute(sql1)
		  gicuitem = rs1(0)
		  
		  sql6 = "INSERT INTO CuDTx7 ([giCuItem]) VALUES(" & gicuitem & ") "
		  conn.Execute(sql6)
			path = ""			
			newsiteid = GetSiteId( insert1 )
			newpath = GetPath( newsiteid, insert1 ,rs("CtRootId") ,rs("CtNodeId"))
			
			
		  'response.write gicuitem
		  sql1i = "INSERT INTO [mGIPcoanew].[dbo].[KnowledgeJigsaw]([gicuitem],[CtRootId],[CtUnitId],[parentIcuitem],[ArticleId],[Status],[orderArticle],[path]) "
		  sql1i = sql1i & "VALUES(" & gicuitem & ", " & request("CtRootId") & ", '" & request("CtRootId") & "', " & Regicuitem & ", "
			sql1i = sql1i & insert1 & ", 'Y', " & orderArticle & ", '" & path & "') "
		  conn.Execute(sql1i)
		end if 
  next
	showDoneBox "編修成功！"
end if

function GetSiteId(id)
		asql = "SELECT siteId FROM CuDTGeneric WHERE icuitem = " & id
		set ars = conn.execute(asql)
		if not ars.eof then
			newsiteid = ars("siteId")
		end if
		ars.close
		set ars = nothing
		GetSiteId = newsiteid
end function

function GetPath( sid, id ,ctRootId, CtNodeId)
	path = ""
	showtype = ""
	xurl = ""
	filedownload = ""
	
	if sid = "3" then					
		topcat = ""
		path = "/knowledge/knowledge_cp.aspx?ArticleId={0}&ArticleType={1}&CategoryId={2}"
		sql = "SELECT topCat FROM CuDTGeneric WHERE iCUItem = " & id
		set prs = conn.execute(sql)
		if not prs.eof then
			topcat = prs("topCat")
		end if
		prs.close
		set prs = nothing
		path = replace(path, "{0}", id)
		path = replace(path, "{1}", "A")
		path = replace(path, "{2}", topcat)
	elseif sid = "2" then
		'---利用id來找目前showtype---		
		sql = "SELECT showType, xURL, fileDownLoad FROM CuDTGeneric WHERE icuitem = " & id
		set prs = conn.execute(sql)
		if not prs.eof then
			showtype = prs("showType")
			xurl = prs("xURL")
			filedownload = prs("fileDownLoad")
		end if
		prs.close
		set prs = nothing
		
		if showtype = "1" then
			nodeid = ""
			mp = ""
			path = "/subject/ct.asp?xItem={0}&ctNode={1}&mp={2}"
			sql = "SELECT CatTreeNode.CtNodeID, CatTreeNode.CtRootID FROM CatTreeNode "
			sql = sql & "INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
			sql = sql & "INNER JOIN CuDTGeneric ON CtUnit.CtUnitID = CuDTGeneric.iCTUnit WHERE CuDTGeneric.icuitem = " & id
						           
			
			set prs = conn.execute(sql)
			if not prs.eof then
				nodeid = prs("CtNodeID")
				mp = prs("CtRootID")
			end if
			prs.close
			set prs = nothing
			path = replace(path, "{0}", id)
			path = replace(path, "{1}", nodeid)
			path = replace(path, "{2}", mp)
		elseif showtype = "2" then
			path = xurl
		elseif showtype = "3" then
			path = "/public/data/" & filedownload
		end if
	elseif sid = "1" then
		sql = "SELECT showType, xURL, fileDownLoad FROM CuDTGeneric WHERE icuitem = " & id
		set prs = conn.execute(sql)
		if not prs.eof then
			showtype = prs("showType")
			xurl = prs("xURL")
			filedownload = prs("fileDownLoad")
		end if
		prs.close
		set prs = nothing		
		if showtype = "1" then 
			nodeid = ""
			path = "/ct.asp?xItem={0}&ctNode={1}&mp=1"
			sql = "SELECT CatTreeNode.CtNodeID FROM CatTreeNode INNER JOIN CtUnit ON CatTreeNode.CtUnitID = CtUnit.CtUnitID "
			sql = sql & "INNER JOIN CuDTGeneric ON CtUnit.CtUnitID = CuDTGeneric.iCTUnit  "
			sql = sql & " WHERE CatTreeNode.ctRootId= '" & ctRootId & "'"
			sql = sql & " and  CatTreeNode.CtNodeId= '" & CtNodeId & "' " 
			sql = sql & " and CuDTGeneric.icuitem = " & id
			
			
			set prs = conn.execute(sql)
			if not prs.eof then
				nodeid = prs("CtNodeID")			
			end if
			prs.close
			set prs = nothing
			path = replace(path, "{0}", id)
			path = replace(path, "{1}", nodeid)		
		elseif showtype = "2" then
			path = xurl
		elseif showtype = "3" then
			path = "/public/data/" & filedownload
		end if
	end if
	GetPath = path
end function

Function CheckExist(articleid)	
	sql3 = "SELECT * FROM [mGIPcoanew].[dbo].[KnowledgeJigsaw] where Status='Y' and [ArticleId]= " & articleid & " AND parenticuitem = " & parenticuitem
	Set RsC = Conn.Execute(sql3)
	If Not RsC.Eof Then
		CheckExist = true
	Else
		CheckExist = false
	End If
	RsC = Empty		
End Function

'response.redirect "subjectPubList.asp?iCUItem=" & Request("iCUItem")
%>
<% Sub showDoneBox(lMsg) %>
  <html>
		<head>
			<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
			<meta name="GENERATOR" content="Hometown Code Generator 1.0">
			<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
			<title>編修表單</title>
    </head>
    <body>
			<script language=vbs>
			  alert("<%=lMsg%>")
              window.location.href="subject_query.asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=request.querystring("gicuitem")%>"
			</script>
    </body>
  </html>
<% End sub %>
