<%@ CodePage = 65001 %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="資料清單"
HTProgCode="GC1AP2"
HTProgPrefix="DsdXML" %>
<!--#INCLUDE FILE="DsdXMLListParam.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htCodeGen.inc" -->
<%
  	set htPageDom = session("codeXMLSpec")
  	set allModel2 = session("codeXMLSpec2")  	    	
  	set refModel = htPageDom.selectSingleNode("//dsTable")
  	set allModel = htPageDom.selectSingleNode("//DataSchemaDef")

if request("submitTask")<>"" then
    SQLUpdate=""
    For Each x In Request.Form
	if left(x,5)="ckbox" and request(x)<>"" then
	    xn=mid(x,6)
	    SQLUpdate=SQLUpdate+"Update CuDTGeneric Set fCTUPublic='Y' where iCUItem=" & request("xphKeyiCUItem"&xn) & ";"
        end if
    Next 
    if SQLUpdate<>"" then conn.execute(SQLUpdate)
%>
	<script language=VBS>
		alert "審核完成！"
		window.navigate "<%=HTProgPrefix%>List.asp?CtRootID=<%=request("CtRootID")%>"
	    </script>
<%	   
	response.end
else
    nowPage=Request.QueryString("nowPage")  '現在頁數
    if nowPage="" then
	xSelect = "htx.*, ghtx.*"
	xFrom = nullText(refModel.selectSingleNode("tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.iCUItem=htx.giCUItem "
	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[showList!='' and refLookup!='']")
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

	for each param in allModel.selectNodes("//dsTable[tableName='CuDTGeneric']/fieldList/field[showList!='' and refLookup!='']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeID='" & param.selectSingleNode("refLookup").text & "'"
		if nullText(param.selectSingleNode("fieldName"))="fCTUPublic" then
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
	    	xFrom = xFrom & " AND " & xAlias & "." & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'"
		xFrom = xFrom & ")"
        ' --- 把 detailRow 裡的 refField 換掉
'        	param.selectSingleNode("fieldName").text = xAFldName
        ' -----------------------------------
	next

	fSql = "SELECT " & xSelect & " FROM " & xFrom 
	fSql = fSql & " WHERE 2=2 "
	if (HTProgRight AND 64) = 0 then
		fSql = fSql & " AND ghtx.iDept=" & pkStr(session("deptID"),"")
	end if
	if session("fCtUnitOnly")="Y" then	fSql = fSql & " AND ghtx.iCtUnit=" & session("CtUnitID") & " "
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
	if nullText(refModel.selectSingleNode("orderList")) <> "" then
		fSql = fSql & " ORDER BY " & refModel.selectSingleNode("orderList").text
	else
		fSql = fSql & " ORDER BY dEditDate DESC"
	end if
	session("GipApproveDsdXMLList")=fSql
    else
    	fSql=session("GipApproveDsdXMLList")
    end if
'response.write fSql
'response.end


if nowPage="" then nowPage=1
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
	    <td class="FormName"><%=HTProgCap%>&nbsp;
	    <font size=2>【主題單元:<%=session("CtUnitName")%> / 單元資料:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】</td>
    <td class="FormLink" valign="top" align=right>
		<%if (HTProgRight and 2)=2 then%>
			<A href="DsdXMLQuery.asp" title="指定查詢條件">查詢</A>
		<%end if%>
		<%if (HTProgRight and 4)=4 then%>
			<A href="DsdXMLAdd.asp" title="新增代碼資料">新增</A>
		<%end if%>
    </td>    
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right colspan="2">
	   <SPAN id=RunJob style="visibility:hidden">
	   	   <input type=button class=cbutton value="審核通過" id=button1 name=button1>
	   </SPAN>    
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
 <INPUT TYPE=hidden name=submitTask value="">
 <INPUT TYPE=hidden name=CtRootID value="<%=request("ItemID")%>">  
  <p align="center">  
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
	<td align=center class=eTableLable width=7%>
	<input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></td>     
     	<td class=eTableLable>預覽</td>
     	<td class=eTableLable>標 題</td>
<%
	for each param in allModel.selectNodes("//fieldList/field[showList!='']")
		response.write "<td class=eTableLable>" & nullText(param.selectSingleNode("fieldLabel")) & "</td>"
	next
    response.write "</tr>"
If not RSreg.eof then   

    for i=1 to PerPageSize
    	xUrl = session("myWWWSiteURL") & "/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("iCuItem")
    	pKey = "iCuItem=" & RSreg("iCuItem")
%>                  
<tr>   
<%if RSreg("fCTUPublic")="P" then%>
<TD align=center><INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y">
<INPUT TYPE=hidden name="xphKeyiCUItem<%=i%>" value="<%=RSreg("iCUItem")%>"></td>   
<%else%>
<TD align=center>&nbsp;</td>   
<%end if%>            
	<TD class=eTableContent><font size=2>
	<A href="<%=xUrl%>" target="_wMof">View</A></TD>
	<TD class=eTableContent><font size=2>
		<A href="DsdXMLEdit.asp?<%=pKey%>">
		<%=RSreg("sTitle")%>
		</A>
	</TD>
<%	
	xrCount = 0
	for each param in allModel.selectNodes("//fieldList/field[showList!='']")
		kf = param.selectSingleNode("fieldName").text
		if nullText(param.selectSingleNode("refLookup")) <> "" then
			xrCount = xrCount + 1
			kf = "xref" & xrCount & kf
		end if
%>
	<TD class=eTableContent><font size=2>
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

      Dim chkCount
      chkCount=0            '記錄checkbox 被勾數

    sub document_onClick           'checkbox 被勾計數
         set sObj=window.event.srcElement
         if sObj.tagName="INPUT" then 
            if sObj.type="checkbox"  then 
                if sObj.checked then 
                   chkCount=chkCount+1
                else
                   chkCount=chkCount-1                
                end if                                          
            end if
         end if
         '
         if chkCount=0 then 
            document.all("RunJob").style.visibility="hidden"
         else
            document.all("RunJob").style.visibility="visible"
         end if
    end sub        
    sub Chkall
       chkCount=0     
       if document.all("ckall").value="全選" then           '全勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=true
                   chkCount=chkCount+1
              end if     
          next                 
          document.all("RunJob").style.visibility="visible"
          document.all("ckall").value="全不選"
      elseif document.all("ckall").value="全不選" then        '全不勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=false
               end if     
          next                 
          document.all("RunJob").style.visibility="hidden"
          document.all("ckall").value="全選"          
       end if
   end sub  
   sub button1_onClick
	reg.submitTask.value="UPDATE"
	reg.submit
   end sub
</script>
