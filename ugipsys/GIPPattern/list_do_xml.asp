<%@ CodePage = 65001 %>
<?xml version="1.0"  encoding="utf-8" ?>
<hpMain>
<!--#Include virtual = "/inc/client.inc" -->
<% 


function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function

	specID=request.querystring("specID")
	set specDom= server.createObject("microsoft.XMLDOM")
	specDom.async = false
	specDom.setProperty("ServerHTTPRequest") = true	
	xv = specDom.load(server.mappath("xmlSpec/"&specID&"List.xml"))
	if specDom.parseError.reason <> "" then 
		Response.Write("specDom parseError on line " &  specDom.parseError.line)
		Response.Write("<BR>Reason: " &  specDom.parseError.reason)
		Response.End()
	end if
	set pageSpecNode=specDom.selectSingleNode("//htPage/pageSpec")
	set resultSetNode=specDom.selectSingleNode("//htPage/resultSet")
	response.write "<contextPath></contextPath>"
	response.write "<headScript></headScript>"
	response.write "<funcName>"&nullText(pageSpecNode.selectSingleNode("pageHead"))&"</funcName>"
	response.write "<navList>"
	for each AnchorNode in pageSpecNode.selectNodes("aidLinkList/Anchor")
	    	response.write "<nav href="""&deAmp(nullText(AnchorNode.selectSingleNode("AnchorURI")))&""">"&nullText(AnchorNode.selectSingleNode("AnchorLabel"))&"</nav>"
	next
	response.write "</navList>"
	response.write "<formName>"&nullText(pageSpecNode.selectSingleNode("pageFunction"))&"</formName>"
	response.write "<mainContent>"
	'----SQL recordsets
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	'----SQL字串
	xSelect=nullText(resultSetNode.selectSingleNode("sql/selectList"))
	xFrom=nullText(resultSetNode.selectSingleNode("sql/fromList"))
	for each colSpecNode in pageSpecNode.selectNodes("detailRow/colSpec")
	    xfieldName=nullText(colSpecNode.selectSingleNode("content/refField"))
	    xinputType=nullText(resultSetNode.selectSingleNode("fieldList/field[fieldName='"&xfieldName&"']/inputType"))
	    xrefLookup=nullText(resultSetNode.selectSingleNode("fieldList/field[fieldName='"&xfieldName&"']/refLookup"))
	    if xrefLookup<>"" and xinputType<>"refCheckbox" and xinputType<>"refCheckboxOther" then    
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeId='" & xrefLookup & "'"
        	SET RSLK=conn.execute(SQL)  
        	xAFldName = xAlias & xfieldName
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = htx." & xfieldName
		if not isNull(RSLK("CodeSrcFld")) then _
	    		xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"	    
	    end if
	next		
	fSql=xSelect&" "&xFrom&" "&nullText(resultSetNode.selectSingleNode("sql/whereList"))
	fSql=fSql&" "&nullText(resultSetNode.selectSingleNode("sql/orderList"))
	'----SQL字串結束

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------

	if Not RSreg.eof then 
	    totRec=RSreg.Recordcount       '總筆數
	    if totRec>0 then 
		PerPageSize=cint(Request.QueryString("pagesize"))
		if PerPageSize <= 0 then  
		    PerPageSize=30  
		end if 	      
	        RSreg.PageSize=PerPageSize       '每頁筆數	
		if cint(nowPage)<1 then 
		    nowPage=1
		elseif cint(nowPage) > RSreg.PageCount then 
		    nowPage=RSreg.PageCount 
		end if 
	        RSreg.AbsolutePage=nowPage	'目前頁數
	        totPage=RSreg.PageCount       '總頁數
	    end if    
	end if
'	response.write "<SQL>"&fSql&"</SQL>"
	response.write "<page>"
	response.write "<hiddenInput><listspec>"&specID&"List.xml</listspec><gipdsd></gipdsd><ictunit></ictunit><kpsn></kpsn></hiddenInput>"
	
	response.write "<totaldata>"&totRec&"</totaldata>"
	response.write "<totalPage>"&totPage&"</totalPage>"
	response.write "<pageSize>"&PerPageSize&"</pageSize>"
	response.write "<curPage>"&nowPage&"</curPage>"
	response.write "</page>"
	response.write "<UIParts></UIParts>"
	response.write "<TopicList>"
	response.write "<TopicName>"&specID&"</TopicName>"
	'----條列表頭
	response.write "<ColumnHead>"
	colSpecCount=0
	for each colSpecNode in pageSpecNode.selectNodes("detailRow/colSpec")
		response.write "<Column id="""&colSpecCount&"""><value>"&nullText(colSpecNode.selectSingleNode("colLabel"))&"</value></Column>"
		colSpecCount=colSpecCount+1
	next
	response.write "</ColumnHead>"
	'----條列表身
	if not RSreg.eof then
	    response.write "<Group id ="""" name="""">"
	    for i=1 to PerPageSize  
	    	response.write "<Article>"
		colSpecCount=0
		xrCount2=0
		'----pKey值
		pKey=""
		for each fieldNode in resultSetNode.selectNodes("fieldList/field[isPrimaryKey='Y']")
		    pKey=pKey&"&amp;" & nullText(fieldNode.selectSingleNode("fieldName")) & "=" & RSreg(nullText(fieldNode.selectSingleNode("fieldName")))
		next	
'		response.write "<pKey>"&pKey&"</pKey>"
		for each colSpecNode in pageSpecNode.selectNodes("detailRow/colSpec")		
		    xfieldName=nullText(colSpecNode.selectSingleNode("content/refField"))
		    xinputType=nullText(resultSetNode.selectSingleNode("fieldList/field[fieldName='"&xfieldName&"']/inputType"))
		    xrefLookup=nullText(resultSetNode.selectSingleNode("fieldList/field[fieldName='"&xfieldName&"']/refLookup"))
		    if xrefLookup<>"" and xinputType<>"refCheckbox" and xinputType<>"refCheckboxOther" then   
			xrCount2 = xrCount2 + 1
			xAlias = "xref" & xrCount2
        		xfieldName = xAlias & xfieldName	    		
		    end if
		    fieldValue=""
		    if not isNull(RSreg(xfieldName)) then fieldValue=RSreg(xfieldName)
		    response.write "<Column id="""&colSpecCount&""">"
		    if nullText(colSpecNode.selectSingleNode("url"))<>"" then _
		    	response.write "<url>"&nullText(colSpecNode.selectSingleNode("url"))&pKey&"</url>"
		    response.write "<value>"&fieldValue&"</value>"
		    response.write "</Column>"
		    colSpecCount=colSpecCount+1
		next		
		response.write "</Article>"
                RSreg.moveNext
                if RSreg.eof then exit for 
	    next
	    response.write "</Group>"
	end if	
	response.write "</TopicList>"	
	response.write "</mainContent>"
%>
</hpMain>
