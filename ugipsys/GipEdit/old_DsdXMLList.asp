<%@ CodePage = 65001 %>
<%Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="資料清單"
HTProgCode="GC1AP1"
HTProgPrefix="old_DsdXML" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
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
 	ListStyle = nullText(htPageDom.selectSingleNode("//ListStyle"))
 	
 
	'======	2006.6.6 by Gary
	'if checkGIPconfig("Discuss") then	
		session("replyID") = 0
		session("replyTitle") = ""
	'end if
	'========	2006.5.25 by Gary  	
if request("submitTask")="DELETE" then

    For Each x In Request.Form
		if left(x,5)="ckbox" and request(x)<>"" then
		    xn=mid(x,6)
			session("BatchDicuitem") = request("xphKeyicuitem"&xn)
			session("BatchDphase") = "edit"
			session("BatchDsubmitTask") = "DELETE"
			postURL = "DsdXMLEdit.asp"
			'======= Server.Execute 不能帶參數
			Server.Execute (postURL) 
        end if
    Next 
	session("BatchDicuitem") = ""
	session("BatchDphase") = ""
	session("BatchDsubmitTask") = ""
%>
	<script language=VBS>
		alert "刪除完成！"
		window.navigate "<%=HTProgPrefix%>List.asp"
	</script>
<%	   
	response.end
elseif request("keep")="" then
   	'========	
	xSelect = " htx.*, ghtx.*"
	if nullText(allModel.selectSingleNode("recordUserRead"))="Y" then
		xSelect = xSelect & ", (SELECT count(*) FROM recordUserRead WHERE rdICuItem=ghtx.iCuitem) AS readCount "
	end if
	for each param in refModel.selectNodes("//field[xSQL!='']")
		xSelect = xSelect & ", " & nullText(param.selectSingleNode("xSQL")) & " AS " _
			& nullText(param.selectSingleNode("fieldName"))
	next
'on error resume next
	xFrom = nullText(refModel.selectSingleNode("tableName")) & " AS htx " _
			& " JOIN CuDtGeneric AS ghtx ON ghtx.icuitem=htx.gicuitem "
	xrCount = 0
	for each param in refModel.selectNodes("fieldList/field[showList!='' and refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		SQL="Select * from CodeMetaDef where codeId=N'" & param.selectSingleNode("refLookup").text & "'"
		if request("debug")<>"" then
			response.write sql & "<HR>"
			response.end
		end if
        SET RSLK=conn.execute(SQL)  
        xAFldName = xAlias & param.selectSingleNode("fieldName").text
'response.write SQL
'response.end
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
	for each param in allModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[showList!='' and refLookup!='']")
		'response.write param.xml & "<HR>" & vbCRLF
	next

	for each param in allModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[showList!='' and refLookup!='' and inputType!='refCheckbox'and inputType!='refCheckboxOther']")
		xrCount = xrCount + 1
		xAlias = "xref" & xrCount
		if nullText(param.selectSingleNode("fieldName"))="fctupublic" and session("CheckYN")="Y" then
			param.selectSingleNode("refLookup").text = "isPublic3"
			SQL="Select * from CodeMetaDef where codeId=N'" & param.selectSingleNode("refLookup").text & "'"		
			param.selectSingleNode("refLookup").text = "isPublic"	
		else
			SQL="Select * from CodeMetaDef where codeId=N'" & param.selectSingleNode("refLookup").text & "'"		
		end if		
		'response.write sql & "<HR>"
        SET RSLK=conn.execute(SQL)  
 		if RSLK.eof then
			response.write sql
			response.end
		end if
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

	fSql = " FROM " & xFrom 
	fSql = fSql & " WHERE 2=2 "
	if checkGIPconfig("Discuss") and nullText(refModel.selectSingleNode("tableName")) = "Discuss" then
	elseif (HTProgRight AND 64) = 0 then
		fSql = fSql & " AND ghtx.idept LIKE '" & session("deptID") & "%' "
	end if
	if session("fCtUnitOnly")="Y" then	fSql = fSql & " AND ghtx.ictunit=" & session("CtUnitID") & " "
	if session("onlyThisNode")="Y" then	fSql = fSql & " AND ghtx.iNode=" & session("ctNodeId") & " "
	if nullText(refModel.selectSingleNode("whereList")) <> "" then _
		fSql = fSql & " AND " & refModel.selectSingleNode("whereList").text
	for each param in allModel.selectNodes("//dsTable[tableName='CuDtGeneric']/fieldList/field[paramKind]")
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
	session("baseSql") = fSql
	session("xSelectSql") = xSelect
	
end if

	fSql = session("baseSql")
	xSelect = session("xSelectSql")
	cSql = "SELECT count(*) " & fSql
	
	if request("shortLongList")<>"" then	session("shortLongList") = request("shortLongList")

nowPage=Request.QueryString("nowPage")  '現在頁數
if nowpage = "" then nowpage = 1
      PerPageSize=Request.QueryString("pagesize")
if perpagesize = "" then perpagesize = 0
      if PerPageSize <= 0 then  PerPageSize=15  
' response.write cSql
' response.end     
      set RSc = conn.execute(cSql)
	  totRec=RSc(0)       '總筆數
'      response.write totRec & "<HR>"
      totPage = int(totRec/PerPageSize+0.999)
'      response.write totPage & "<HR>"
      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) >totPage then 
         nowPage=totPage 
      end if            	

	if nullText(allModel.selectSingleNode("//showClientSqlOrderBy")) <> "" then
		fSql = fSql & " " & nullText(allModel.selectSingleNode("//showClientSqlOrderBy"))
	else
		fSql = fSql & " ORDER BY ghtx.xpostDate DESC"
	end if
	fSql = "SELECT TOP " & nowPage*PerPageSize & xSelect & fSql
'	response.write fSql & "<HR>"
'	response.end


 Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,3,1
Set RSreg = Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------

'response.write fSql
if Not RSreg.eof then

   if totRec>0 then 
      RSreg.PageSize=PerPageSize       '每頁筆數
      RSreg.AbsolutePage=nowPage
      strSql=server.URLEncode(fSql)
   end if    
end if   
%>

<body>
<div id="FuncName">
	<h1><%=Title%></h1><font size=2>【目錄樹節點: <%=session("catName")%>】</font>
	<div id="Nav">
		<%if (HTProgRight and 4)=5 then%>
			<A href="UserCatNodeAdd.asp?phase=add" title="新增分類節點">新增分類節點</A>
		<%end if%>
		<%if (HTProgRight and 2)=2 and 0 then%>
			<A href="DsdXMLQuery.asp" title="指定查詢條件">查詢</A>
		<%end if%>
		<%if (HTProgRight and 4)=4 and 0 then%>
			<A href="DsdXMLAdd.asp?phase=add" title="新增資料">新增</A>
		<%end if%>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">
	<%=HTProgCap%>&nbsp;
	<font size=2>【主題單元:<%=session("CtUnitName")%> / 單元資料:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>】
</div>
<!--  條列頁簡易查詢功能  -->
<% IF checkGIPconfig("AttachmentType") then %>
<form method="POST" name="reg" action="DsDXMLList.asp">
<INPUT TYPE=hidden name=submitTask value="">
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD>簡易查詢</TD></TR>
<TR><TD class="eTableLable">標 題： <input name="htx_stitle" size="30"><input name="Submit" type="submit" class="Button" value="查詢"/>
</TD></TR>
</TABLE>
</form>     
<% END IF %>

<Form id="Form2" name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
 <INPUT TYPE=hidden name=submitTask value="">
	<!-- 分頁 -->
	<div id="Page">
    <SPAN id=RunJob style="visibility:hidden">
	   <input type=button class=cbutton value="刪除" id=button1 name=button1>
	</SPAN>			
      <% if cint(nowPage) <>1 then %>
			<img src="/images/arrow_previous.gif" alt="上一頁">       		
			<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> ，
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
        <% if cint(nowPage)<>totPage then %> 
            ，<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁
            <img src="/images/arrow_next.gif" alt="下一頁"></a> 
        <%end if%>     
<%
	if nullText(allModel.selectSingleNode("//field[longListOnly='Y']/longListOnly"))="Y" then
		response.write "<BUTTON onClick=""VBS: shortLongList 'short'"">簡目式</BUTTON>"
		response.write "<BUTTON onClick=""VBS: shortLongList 'long'"">詳目式</BUTTON>"
	end if
%>
	</div>

	<!-- 列表 -->
	<table cellspacing="0" id="ListTable">
		<tr>
<%	
	'======	2006.5.24 by Gary
	if session("catName")="GaryTest" then
		%>
		<th class="First" scope="col" width=7%><input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></th>
		<%
	end if
	'======	2006.5.10 by Gary
	if checkGIPconfig("RSSandQuery") and nullText(refModel.selectSingleNode("tableName")) = "RSSRead" then 
		%>
		<th class="First" scope="col">執行</th>
		<%
	else
		%>
		<th class="First" scope="col">預覽</th>
		<%
	end if
	'======		
%>			<!--<th class="First" scope="col">預覽</th>   -->
			<th><p align="center">標 題</p></th>
<%
	lorsCheck = "//fieldList/field[showList!='' and fieldName!='stitle' and (not (longListOnly))]"
	if session("shortLongList")="long" then		lorsCheck = "//fieldList/field[showList!='' and fieldName!='stitle']"
	
	for each param in allModel.selectNodes(lorsCheck)
		response.write "<th><p align=""center"">" & nullText(param.selectSingleNode("fieldLabel")) & "</th>"
	next
 	'======	2006.6.6 by Gary
	if checkGIPconfig("Discuss") and nullText(refModel.selectSingleNode("tableName")) = "Discuss" then
	
		%>
		<th><p align="center">回覆數</th>
		<%
	end if
	'======
	'======	2006.7.10 by Chris
	if nullText(allModel.selectSingleNode("recordUserRead"))="Y" then
	
		%>
		<th><p align="center">讀取數</th>
		<%
	end if
	'======
   response.write "</tr>"
If not RSreg.eof then   

    for i=1 to PerPageSize
		'判斷是不是主題館 是的話網址加上subject
		if session("pvXdmp") <> "1" then
			if instr(session("myWWWSiteURL") , "subject") > 0 then
				xUrl = session("myWWWSiteURL") & "/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("icuitem")
			else
				xUrl = session("myWWWSiteURL") & "/subject/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("icuitem")
			end if
		else 
			xUrl = session("myWWWSiteURL") & "/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("icuitem")
		end if
'    	xUrl = "http://www.boaf.gov.tw/boafwww/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("icuitem")
    	pKey = "icuitem=" & RSreg("icuitem")
%>                  
<tr>                  
<%	'======	2006.5.24 by Gary
	if session("catName")="GaryTest" then
		%>
		<TD align=center><INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y">
		<INPUT TYPE=hidden name="xphKeyicuitem<%=i%>" value="<%=RSreg("icuitem")%>"></td> 
		<%
	end if
	'======		
%>	   
	<TD class=eTableContent>
<%	'======	2006.5.10 by Gary
	if checkGIPconfig("RSSandQuery") and nullText(refModel.selectSingleNode("tableName")) = "RSSRead" then 
		xUrl = "/site/" & session("mySiteID") & "/wsxd/ws_RSSRead.asp?gicuitem="  & RSreg("icuitem")
		%>
		<A href="<%=xUrl%>" target="_wMof">Run</A>
		<%
	else
		%>
		<A href="<%=xUrl%>" target="_wMof">View</A>
		<%
	end if
	'======		
%>	
	</TD>
	<TD class=eTableContent>
<%	'======	2006.5.10, 2006.6.6, 2006.6.13	by Gary	
	if ListStyle <> "View" then 
		ListStyle = "Edit"
	end if
	'======	
		%>
		<A href="<%=HTProgPrefix%><%=ListStyle%>.asp?<%=pKey%>&phase=edit&keep=<%=request("keep")%>&nowpage=<%=nowpage%>&pagesize=<%=perpagesize%>">
		<%=RSregFieldShowOut(htPageDom.selectSingleNode("//fieldList/field[fieldName='stitle']"), "sTitle")%>
		</A>
	</TD>
<%	
	xrCount = 0
	for each param in allModel.selectNodes("//fieldList/field[showList!='' and fieldName!='stitle']")
		kf = param.selectSingleNode("fieldName").text
		if nullText(param.selectSingleNode("refLookup")) <> "" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckbox" _
			AND nullText(param.selectSingleNode("inputType"))<>"refCheckboxOther" then
			xrCount = xrCount + 1
			kf = "xref" & xrCount & kf
		end if
		if session("shortLongList")<>"long" AND nullText(param.selectSingleNode("longListOnly"))="Y" then
		else
%>
	<TD class=eTableContent>
<%	if nullText(param.selectSingleNode("xURL")) <> "" then
		response.write "<A href=""" & nullText(param.selectSingleNode("xURL")) _
			& "&xNode=" & session("ctNodeId") & "&" & pKey & """>" & RSregFieldShowOut(param, kf) & "</A>"
	else
		response.write RSregFieldShowOut(param, kf)
	end if %>
</font></td>
<%
		end if
	next
	'======	2006.6.6 by Gary
	if checkGIPconfig("Discuss") and nullText(refModel.selectSingleNode("tableName")) = "Discuss" then 
		sql_reply="SELECT COUNT(gicuitem) FROM Discuss WHERE reply=" & RSreg("gicuitem")
		set rs_reply=conn.Execute(sql_reply)
		%>
		<TD><%=rs_reply(0)%></TD>
		<%
	end if
	'======			
	'======	2006.7.10 by Chris 
	if nullText(allModel.selectSingleNode("recordUserRead"))="Y" then
		%>
		<TD align="center"><A href="userReadList.asp?iCuItem=<%=RSreg("iCuItem")%>"><%=RSreg("readCount")%></A></TD>
		<%
	end if
	'======			
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

<%	function RSregFieldShowOut(xparam, xkf)
		xoStr = RSreg(xkf)
		xoTag = nullText(xparam.selectSingleNode("optionTag/tag"))
'		response.write xoTag & xkf
		if xoTag<>"" then
			if xmlValueCompare(xparam.selectSingleNode("optionTag/condition")) then
				xoStr = "<" & xoTag & ">" & xoStr & "</" & xoTag & ">"
			end if
		end if
		RSregFieldShowOut = xoStr
	end function 
	
	function xmlValueCompare(ckNode)
		xmlValueCompate = false
		select case nullText(ckNode.selectSingleNode("ckLValue/type"))
			case "Data"
				lValue = RSreg(nullText(ckNode.selectSingleNode("ckLValue/ckFieldName")))
			case "Value"
				lValue = nullText(ckNode.selectSingleNode("ckLValue/value"))
			case else
				exit function
		end select
		select case nullText(ckNode.selectSingleNode("ckRValue/type"))
			case "Data"
				rValue = RSreg(nullText(ckNode.selectSingleNode("ckRValue/ckFieldName")))
			case "Value"
				rValue = nullText(ckNode.selectSingleNode("ckRValue/value"))
			case else
				exit function
		end select
		select case nullText(ckNode.selectSingleNode("ckCondition"))
			case "ngt"
				if cSng(lValue) > cSng(rValue) then xmlValueCompare = true
'				response.write lValue & ">" & rValue
'				response.write (cSng(lValue) > cSng(rValue))
			case "nlt"
				if cSng(lValue) < cSng(rValue) then xmlValueCompare = true
			case "nge"
				if cSng(lValue) >= cSng(rValue) then xmlValueCompare = true
			case "nle"
				if cSng(lValue) <= cSng(rValue) then xmlValueCompare = true
			case "gt"
				if lValue > rValue then xmlValueCompare = true
				response.write lValue & ">" & rValue
				response.write (lValue > rValue)
			case "lt"
				if lValue < rValue then xmlValueCompare = true
				response.write lValue & "<" & rValue
			case "ge"
				if lValue >= rValue then xmlValueCompare = true
			case "le"
				if lValue <= rValue then xmlValueCompare = true
			case else
				response.write "ckCondition not match!!!"
				exit function
		end select
	end function
%>

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=" & newPage & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>" & "&pagesize=" & newPerPage                    
     end sub 

     sub shortLongList(sorl)
           document.location.href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>" _
			& "&shortLongList=" & sorl
     end sub 

     sub setpKey(xv)
     	gpKey = xv
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
   	'if document.all("sfx_fctupublic").value="" then
   	'	alert "請點選是否公開狀態欄位!"
   	'	document.all("sfx_fctupublic").focus
   	'	exit sub
   	'end if
   	'xPos=instr(document.all("sfx_fctupublic").value,"--")
   	'fctupublicValue=Left(document.all("sfx_fctupublic").value,xPos-1)
   	'fctupublicDisplay=mid(document.all("sfx_fctupublic").value,xPos+2)
        chky=msgbox("注意！"& vbcrlf & vbcrlf &"　您確定要刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
              'document.all("fctupublic").value=fctupublicValue
	      reg.submitTask.value = "DELETE"
	      reg.Submit
       	end If   	
   end sub
</script>
