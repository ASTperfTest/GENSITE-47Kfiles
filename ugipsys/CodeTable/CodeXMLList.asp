<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="代碼維護"
HTProgFunc="代碼清單"
HTProgCode="Pn90M02"
HTProgPrefix="CodeXML" %>
<!--#INCLUDE FILE="codeXMLListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htCodeGen.inc" -->
<%
  	set htPageDom = session("codeXMLSpec")
  	set refModel = htPageDom.selectSingleNode("//dsTable")

 Set RSreg = Server.CreateObject("ADODB.RecordSet")
 fSql=Request.QueryString("strSql")

if fSql="" then
	xSelect = "htx.*"
	xFrom = nullText(refModel.selectSingleNode("tableName")) & " AS htx "
	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[showList='Y' and refLookup!='']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeId=N'" & param.selectSingleNode("refLookup").text & "'"
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

	fSql = "SELECT " & xSelect & " FROM " & xFrom 
	fSql = fSql & " WHERE 1=1 " 
	if nullText(refModel.selectSingleNode("whereList")) <> "" then _
		fSql = fSql & " AND " & refModel.selectSingleNode("whereList").text

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
	if nullText(refModel.selectSingleNode("orderList")) <> "" then
		fSql = fSql & " ORDER BY " & refModel.selectSingleNode("orderList").text
	end if
'	response.write fSql & "<HR>"
'	response.end
end if
'response.write fSql
'response.end

nowPage=Request.QueryString("nowPage")  '現在頁數


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
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
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;
	    <font size=2>【代碼ID:<%=session("CodeID")%>/代碼名稱:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right>
		<%if (HTProgRight and 2)=2 then%>
			<A href="CodeXMLQuery.asp" title="指定查詢條件">查詢</A>
		<%end if%>
		<%if (HTProgRight and 4)=4 then%>
			<A href="CodeXMLAdd.asp" title="新增代碼資料">新增</A>
		<%end if%>
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
  <p align="center">  
<%If not RSreg.eof then%>     
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
       <% if cint(nowPage) <>1 then %>             
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
             <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
        </select>     
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
<%
	for each param in refModel.selectNodes("fieldList/field[showList='Y']")
		response.write "<td class=eTableLable>" & nullText(param.selectSingleNode("fieldLabel")) & "</td>"
	next
    response.write "</tr>"

    for i=1 to PerPageSize                  
%>                  
<tr>                  
<%pKey = ""
	for each param in refModel.selectNodes("fieldList/field[isPrimaryKey='Y']")
		kf = param.selectSingleNode("fieldName").text
		pKey = pKey & "&" & kf & "=" & RSreg(kf)
	next
	if pKey<>"" then  pKey = mid(pKey,2)

	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[showList='Y']")
		kf = param.selectSingleNode("fieldName").text
		if nullText(param.selectSingleNode("refLookup")) <> "" then
			xrCount = xrCount + 1
			kf = "xref" & xrCount & kf
		end if
%>
	<TD class=eTableContent><font size=2>
	<% if nullText(param.selectSingleNode("isPrimaryKey")) = "Y" then %>
		<A href="CodeXMLEdit.asp?<%=pKey%>">
		<%=RSreg(kf)%>
		</A>
	<% else %>
		<%=RSreg(kf)%>
	<% end if %>
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
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
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
