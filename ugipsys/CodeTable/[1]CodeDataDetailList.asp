<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="代碼維護"
HTProgCode="Pn90M02"
HTProgPrefix="CodeDataDetail" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<%	session("SortFlag")=false
	session("codeId")=request.querystring("codeId")
	
	SQL="Select * from CodeMetaDef where codeId=" & pkStr(request.querystring("codeID"),"")
	SET RSCode=conn.execute(SQL)

	 session("CodeName")=RSCode("CodeName")	
	if not ISNULL(RSCode("CodeSortFld")) then session("SortFlag")=true
	if RSCode("CodeXMLSpec") <> "" then 
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath(".") & "\XMLSpec\" & RSCode("CodeXMLSpec") & ".xml"
		response.write LoadXML & "<HR>"
		xv = htPageDom.load(LoadXML)
'		response.write xv & "<HR>"
  		if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  		end if
  	set session("codeXMLSpec") = htPageDom
  	
%>
<script language=vbs>
	window.navigate "CodeXMLList.asp?codeID=<%=session("codeId")%>"
</script>
<%  	response.end
	end if
Set RSreg = Server.CreateObject("ADODB.RecordSet")

'fSql=Request.QueryString("strSql")  
 If Request.QueryString("nowPage") = "" Then
    session("strSql") = ""
  End If
  fSql = session("strSql")
fsql2=""
if fSql="" or request.querystring("T")="Query" then
    if isNull(RSCode("CodeSrcItem")) then
	if isNull(RSCode("CodeSortFld")) then    
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & " From " & RSCode("CodeTblName")
	else
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & " From " & RSCode("CodeTblName") & " Order By " & RSCode("CodeSortFld")	
	end if		
    else
	if isNull(RSCode("CodeSortFld")) then
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & " From " & RSCode("CodeTblName") & " where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "'"	
	else    
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & " From " & RSCode("CodeTblName") & " where " & RSCode("CodeSrcFld") & "=N'" & RSCode("CodeSrcItem") & "' Order By " & RSCode("CodeSortFld")
	end if
    end if
    if request.querystring("T")="Query" then
    	if request("tfx_CodeValue")<>"" then
    		fsql2=RSCode("CodeValueFld") & " LIKE N'%" & request("tfx_CodeValue") & "%' "
    	end if
    	if request("tfx_CodeDisplay")<>"" then
    	   if fsql2<>"" then
    		fsql2=fsql2 & " AND " & RSCode("CodeDisplayFld") & " LIKE N'%" & request("tfx_CodeDisplay") & "%'"    	
    	   else
    		fsql2=RSCode("CodeDisplayFld") & " LIKE N'%" & request("tfx_CodeDisplay") & "%' "    	    	
    	   end if
    	end if  
    	if request("tfx_CodeSort")<>"" then
    	   if fsql2<>"" then
    		fsql2=fsql2 & " AND " & RSCode("CodeSortFld") & " LIKE N'%" & request("tfx_CodeSort") & "%'"    	
    	   else
    		fsql2=RSCode("CodeSortFld") & " LIKE N'%" & request("tfx_CodeSort") & "%' "    	    	
    	   end if
    	end if    	
    	if fsql2<>"" then 
		pos1=instr(fSql,"where")
        	pos2=instr(fSql,"Order By")
		if pos1=0 and pos2=0 then
	    		fSql=fSql & " where " & fsql2
	    	elseif pos1=0 and pos2<>0 then
		        fSql=left(fSql,pos2-1) & "where " & fsql2 & mid(fSql,pos2-1)
    		elseif pos1<>0 and pos2=0 then
	    		fSql=fSql & " AND " & fsql2 
    		else
	        	fSql=left(fSql,pos2-1) & " AND " & fsql2 & mid(fSql,pos2-1)
    		end if   	
    	end if
    end if    
    if instr(fSql,"Order By")=0 then
		fsql = fsql & " Order By " & RSCode("CodeSortFld")  
    end if
end if
'response.write fsql & "<br>"
nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------

%>
<HTML>
<head>
<title>查詢結果清單</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<Form name=reg method="GET" action=<%=HTprogPrefix%>List.asp>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%> </td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext">【代碼ID:<%=session("codeId")%>/代碼名稱:<%=session("CodeName")%>】</td>
    <td class="FormLink" valign="top">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="CodeDataQuery.asp?codeID=<%=request.querystring("codeId")%>&CodeName=<%=session("CodeName")%>">查詢</a><% End IF %>&nbsp;
      <% if (HTProgRight and 4)=4 then %><a href="CodeDataAdd.asp">新增</a><% End IF %>&nbsp;   
<% 	if checkGIPconfig("codeML") then 
		if RSCode("isML")="Y" then
			response.write "多國語言：<SELECT name=""changeML"" size=""1"">" & vbCRLF
			xSQL = "SELECT * FROM CodeMain WHERE codeMetaID=N'sysML' ORDER BY mSortValue"
			set RSselect = conn.execute(xSQL)
			while not RSselect.eof
				response.write "<OPTION VALUE=" & RSselect("mCode") & ">" & RSselect("mValue") & "</OPTION>"
				RSselect.moveNext
			wend
			response.write "</SELECT>"
		end if
	end if
%></td> 
  </tr>  
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="80%" height=220 valign=top colspan="2">
    
<center>
<font size="2" color="#0000FF">&nbsp;       
<%if RSreg.eof  then
	response.write "<font color=red size=3>====查無資料====</font>"
else

totRec=RSreg.Recordcount       '總筆數

PerPageSize=cint(Request.QueryString("pagesize"))

if PerPageSize <= 0 then  
   PerPageSize=15  
end if 

RSreg.PageSize=PerPageSize       '每頁筆數

if cint(nowPage)<1 then 
   nowPage=1
elseif cint(nowPage) > RSreg.PageCount then 
   nowPage=RSreg.PageCount 
end if            	

RSreg.AbsolutePage=nowPage
totPage=RSreg.PageCount       '總頁數

'strSql=server.URLEncode(fSql)  
session("strSql") = fSql
%>   
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>頁|                      
        共<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
    頁</font>    
        
       <% if nowPage <>1 then %>             
            |<a href="<%=HTProgPrefix%>List.asp?codeID=<%=session("codeID")%>&codeName=<%=session("codeName")%>&nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if nowPage<>RSreg.PageCount then %> 
            |<a href="<%=HTProgPrefix%>List.asp?codeID=<%=session("codeID")%>&codeName=<%=session("codeName")%>&nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>             
        </select>     
     </font>     
</center>
<CENTER>
<TABLE width=95% cellspacing="1" cellpadding="0" class=bg>                   
<tr align=left>    
	<td width=25% align=center class=lightbluetable><%=RSCode("CodeValueFldName")%>(Value)</td>
	<td width=25% align=center class=lightbluetable><%=RSCode("CodeDisplayFldName")%>(Display)</td>
	<%if not isNull(RSCode("CodeSortFld")) then%>
	<td width=25% align=center class=lightbluetable><%=RSCode("CodeSortFldName")%>(SORT)</td>
	<%end if%>
</tr>	                                
  <%                  
   for i=1 to PerPageSize%>                   
<tr>
	<%if not isNull(RSCode("CodeSortFld")) then%>	                  
	<TD class=whitetablebg><p align=center><font size=2><a href="CodeDataEdit.asp?codeID=<% =request.querystring("codeID") %>&value=<%=RSreg(0)%>"><%=RSreg(0)%></A></FONT></TD>
	<%else%>
	<TD class=whitetablebg><p align=center><font size=2><a href="CodeDataEdit.asp?codeID=<% =request.querystring("codeID") %>&value=<%=RSreg(0)%>"><%=RSreg(0)%></A></FONT></TD>	
	<%end if%>
	<TD class=whitetablebg><font size=2>　<%=RSreg(1)%></FONT></TD>
	<%if not isNull(RSCode("CodeSortFld")) then%>	
	<TD class=whitetablebg><font size=2>　<%=RSreg(2)%></FONT></TD>	
	<%end if%>
</tr>
    <% RSreg.moveNext   
     if RSreg.EOF then exit for       
  next%>
    </td>
  </tr>  
  </table> 
</table>
</CENTER>
</form>
</BODY>
</HTML>
<%end if%>
<script language=VBScript>	
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?codeID=<%=session("codeID")%>&codeName=<%=session("codeName")%>&nowPage=" & newPage & "&strSql=<%= strSql%>" & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?codeID=<%=session("codeID")%>&codeName=<%=session("codeName")%>&nowPage=<%=nowPage%>" & "&strSql=<%= strSql%>" & "&pagesize=" & newPerPage                    
     end sub 
     
     sub changeML_onChange
'     	msgBox reg.changeML.value
     	document.location.href="ML<%=HTProgPrefix%>List.asp?codeID=<%=session("codeID")%>&codeName=<%=session("codeName")%>&ML=" & reg.changeML.value
     end sub
</script>
