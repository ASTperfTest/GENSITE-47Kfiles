<%@ CodePage = 65001 %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="資料清單"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#INCLUDE FILE="DsdXMLListParam.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
'	if not isNull(xNode) then
'	  if isObject(xNode) then
'		nullText = 
'	  end if
'	else
'		nullText = "aaa"
'	end if
end function
  	set htPageDom = session("codeXMLSpec")
  	set allModel2 = session("codeXMLSpec2")  	    	
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")
  	
 fSql=Request.QueryString("strSql")
 
if fSql="" then
	xSelect = "htx.*, ghtx.*"
	xFrom = nullText(refModel.selectSingleNode("tableName")) & " AS htx " _
			& " JOIN CuDTGeneric AS ghtx ON ghtx.iCUItem=htx.giCuItem "
	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[showList!='' and refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID=N'" & param.selectSingleNode("refLookup").text & "'"
        SET RSLK=conn.execute(SQL)  
        xAFldName = xAlias & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = htx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
        ' --- 把 detailRow 裡的 refField 換掉
'        	param.selectSingleNode("fieldName").text = xAFldName
        ' -----------------------------------
	next

	for each param in allModel.selectNodes("//dsTable[tableName='CuDTGeneric']/fieldList/field[showList!='' and refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		if nullText(param.selectSingleNode("fieldName"))="fCTUPublic" and session("CheckYN")="Y" then
			param.selectSingleNode("refLookup").text = "isPublic3"
			SQL="Select * from CodeMetaDef where codeID=N'" & param.selectSingleNode("refLookup").text & "'"		
			param.selectSingleNode("refLookup").text = "isPublic"	
		else
			SQL="Select * from CodeMetaDef where codeID=N'" & param.selectSingleNode("refLookup").text & "'"		
		end if		
        SET RSLK=conn.execute(SQL)  
        xAFldName = xAlias & param.selectSingleNode("fieldName").text
		xSelect = xSelect & ", " & xAlias & "." & RSLK("CodeDisplayFld") & " AS " & xAFldName
		xFrom = "(" & xFrom & " LEFT JOIN " & RSLK("CodeTblName") & " AS " & xAlias & " ON " _
			& xAlias & "." & RSLK("CodeValueFld") & " = ghtx." & param.selectSingleNode("fieldName").text
		if not isNull(RSLK("CodeSrcFld")) then _
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "=N'" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
        ' --- 把 detailRow 裡的 refField 換掉
'        	param.selectSingleNode("fieldName").text = xAFldName
        ' -----------------------------------
	next

	fSql = "SELECT " & xSelect & " FROM " & xFrom 
	fSql = fSql & " WHERE 2=2 "
	if (HTProgRight AND 64) = 0 then
		fSql = fSql & " AND ghtx.iDept LIKE '" & session("deptID") & "%' "
	end if
	if session("fCtUnitOnly")="Y" then	fSql = fSql & " AND ghtx.iCTUnit=" & session("CtUnitID") & " "
	if nullText(refModel.selectSingleNode("whereList")) <> "" then _
		fSql = fSql & " AND " & refModel.selectSingleNode("whereList").text
	for each param in allModel.selectNodes("//dsTable[tableName='CuDTGeneric']/fieldList/field[paramKind]")
	  paramKind = nullText(param.selectSingleNode("paramKind"))
	  paramCode = nullText(param.selectSingleNode("fieldName"))
	  paramKindPad = ""
	  if paramKind = "range" then 	paramKindPad = "S"
	  if request.form("htx_" & paramCode & paramKindPad) <> "" then
		select case paramKind
		  case "range"
			rangeS = request.form("htx_" & paramCode & "S")
			rangeE = request.form("htx_" & paramCode & "E")
			if rangeE = "" then	rangeE=rangeS
			whereCondition = replace("ghtx." & paramCode & " BETWEEN '{0}' and '{1}'", _
				"{0}", rangeS)
			whereCondition = replace(whereCondition, "{1}", rangeE)
		  case "value"
			whereCondition = replace("ghtx." & paramCode & " = {0}", "{0}", _
				pkStr(request.form("htx_" & paramCode),""))
		  case else		'-- LIKE
			whereCondition = replace("ghtx." & paramCode & " LIKE {0}", "{0}", _
				pkStr("%"&request.form("htx_" & paramCode)&"%",""))
		end select
		fSql = fSql & " AND " & whereCondition
	  end if
	next
	for each param in refModel.selectNodes("fieldList/field[paramKind]")
	  paramKind = nullText(param.selectSingleNode("paramKind"))
	  paramCode = nullText(param.selectSingleNode("fieldName"))
	  paramKindPad = ""
	  if paramKind = "range" then 	paramKindPad = "S"
	  if request.form("htx_" & paramCode & paramKindPad) <> "" then
		select case paramKind
		  case "range"
			rangeS = request.form("htx_" & paramCode & "S")
			rangeE = request.form("htx_" & paramCode & "E")
			if rangeE = "" then	rangeE=rangeS
			whereCondition = replace("htx." & paramCode & " BETWEEN '{0}' and '{1}'", _
				"{0}", rangeS)
			whereCondition = replace(whereCondition, "{1}", rangeE)
		  case "value"
			whereCondition = replace("htx." & paramCode & " = {0}", "{0}", _
				pkStr(request.form("htx_" & paramCode),""))
		  case else		'-- LIKE
			whereCondition = replace("htx." & paramCode & " LIKE {0}", "{0}", _
				pkStr("%"&request.form("htx_" & paramCode)&"%",""))
		end select
		fSql = fSql & " AND " & whereCondition
	  end if
	next
	xpCondition
	if nullText(allModel.selectSingleNode("//showClientSqlOrderBy")) <> "" then
		fSql = fSql & " " & nullText(allModel.selectSingleNode("//showClientSqlOrderBy"))
	else
		fSql = fSql & " ORDER BY ghtx.xPostDate DESC"
	end if
'	response.write fSql & "<HR>"
'	response.end
end if
'response.write fSql
'response.end

nowPage=Request.QueryString("nowPage")  '現在頁數
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=15  
      end if 

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>

<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
		<%if (HTProgRight and 2)=2 then%>
			<A href="DsdXMLQuery.asp" title="指定查詢條件">查詢</A>
		<%end if%>
		<%if (HTProgRight and 4)=4 then%>
			<A href="DsdXMLAdd.asp" title="新增代碼資料">新增</A>
		<%end if%>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	<%=HTProgCap%>&nbsp;
	<font size=2>【主題單元:<%=session("CtUnitName")%> / 單元資料:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】
</div>

<Form id="Form2" name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
	<!-- 分頁 -->
	<div id="Page">
       <% if cint(nowPage) <>1 then %>
			<img src="/images/arrow_previous.gif" alt="上一頁">       		
			<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> ，
       <%end if%>      
		共<em><%=totRec%></em>筆資料，每頁顯示
       <select id=PerPage size="1" style="color:#FF0000" class="select">            
             <option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
        </select>     
		筆，目前在第
         <select id=GoPage size="1" style="color:#FF0000" class="select">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
		頁	
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            ，<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁
            <img src="/images/arrow_next.gif" alt="下一頁"></a> 
        <%end if%>     
	</div>

	<!-- 列表 -->
	<table cellspacing="0" id="ListTable">
		<tr>
			<th class="First" scope="col">預覽</th>
			<th><p align="center">標 題</p></th>
<%
	for each param in allModel.selectNodes("//fieldList/field[showList!='']")
		response.write "<th><p align=""center"">" & nullText(param.selectSingleNode("fieldLabel")) & "</th>"
	next
    response.write "</tr>"
If not RSreg.eof then   

    for i=1 to PerPageSize
    	xUrl = session("myWWWSiteURL") & "/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("iCuItem")
'    	xUrl = "http://www.boaf.gov.tw/boafwww/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("iCuItem")
    	pKey = "iCuItem=" & RSreg("iCuItem")
%>                  
<tr>                  
	<TD class=eTableContent>
	<A href="<%=xUrl%>" target="_wMof">View</A></TD>
	<TD class=eTableContent>
		<A href="DsdXMLEdit.asp?<%=pKey%>">
		<%=RSreg("sTitle")%>
		</A>
	</TD>
<%	
	xrCount = 0
	for each param in allModel.selectNodes("//fieldList/field[showList!='']")
		kf = param.selectSingleNode("fieldName").text
		if nullText(param.selectSingleNode("refLookup")) <> "" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckbox" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckboxOther" then
			xrCount = xrCount + 1
			kf = "xref" & xrCount & kf
		end if
%>
	<TD class=eTableContent>
		<%=RSreg(kf)%>
</font></td>
<%
	next
%>
    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
    </TABLE>
       <!-- 程式結束 ---------------------------------->  
     </form>  
</body>
</html>                                 
<%else%>
      <script language=vbs>
           msgbox "找不到資料, 請重設查詢條件!"
'	       window.history.back
      </script>
<%end if%>

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

</script>
