<%@ CodePage = 65001 %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="資料清單"
HTProgCode="PEDIA01"
HTProgPrefix="Pedia" 

Dim iCtUnitId : iCtUnitId = session("PediaUnitId")

%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title></title>
</head>

<%
	
	if request("submitTask")="DELETE" then

	elseif request("submitTask")= "QUERY" then
	
		cSql = "SELECT COUNT(*) FROM Pedia INNER JOIN CuDtGeneric ON CuDtGeneric.icuitem = Pedia.gicuitem "
		cSql = cSql & "LEFT JOIN Member ON Pedia.memberId = Member.account "
		cSql = cSql & "WHERE CuDtGeneric.ictunit = " & iCtUnitId 
		if request("sTitle") <> "" then cSql = cSql & " AND sTitle LIKE '%" & request("sTitle") & "%' "
		if request("xPostDateS") <> "" then cSql = cSql & " AND commendTime >= '" & request("xPostDateS") & " 00:00:00' "
		if request("xPostDateE") <> "" then cSql = cSql & " AND commendTime <= '" & request("xPostDateE") & " 23:59:59' "
		if request("fCTUPublic") <> "" then cSql = cSql & " AND fCTUPublic = '" & request("fCTUPublic") & "' "
		if request("xStatus") <> "" then cSql = cSql & " AND xStatus = '" & request("xStatus") & "' "
		if request("memberId") <> "" then 
			cSql = cSql & " AND ( memberId LIKE '%" & request("memberId") & "%' OR realname LIKE '" & request("memberId") & "' "
			cSql = cSql & " OR nickname LIKE '%" & request("memberId") & "%') "
		end if		
		
		fSql = "CuDTGeneric.*, Pedia.*, CodeMain_1.mValue AS isPublicValue, CodeMain.mValue AS isOpenValue, CodeMain_2.mValue AS cataValue, "
		fSql = fSql & "Member.realname, Member.nickname "
		fSql = fSql & "FROM CuDTGeneric LEFT OUTER JOIN CodeMain ON CuDTGeneric.vGroup = CodeMain.mCode "
		fSql = fSql & "AND CodeMain.codeMetaID = 'boolYN' LEFT OUTER JOIN CodeMain AS CodeMain_2 ON CuDTGeneric.topCat = CodeMain_2.mCode "
		fSql = fSql & "AND CodeMain_2.codeMetaID = 'pediacata' LEFT OUTER JOIN CodeMain AS CodeMain_1 "
		fSql = fSql & "ON CuDTGeneric.fCTUPublic = CodeMain_1.mCode AND CodeMain_1.codeMetaID = 'isPublic' "
		fSql = fSql & "INNER JOIN Pedia ON CuDTGeneric.icuitem = Pedia.gicuitem "
		fSql = fSql & "LEFT OUTER JOIN Member ON Pedia.memberId = Member.account "
		fSql = fSql & "WHERE CuDTGeneric.iCTUnit = " & iCtUnitId
		if request("sTitle") <> "" then fSql = fSql & " AND sTitle LIKE '%" & request("sTitle") & "%' "
		if request("xPostDateS") <> "" then fSql = fSql & " AND commendTime >= '" & request("xPostDateS") & " 00:00:00' "
		if request("xPostDateE") <> "" then fSql = fSql & " AND commendTime <= '" & request("xPostDateE") & " 23:59:59' "
		if request("fCTUPublic") <> "" then fSql = fSql & " AND fCTUPublic = '" & request("fCTUPublic") & "' "
		if request("xStatus") <> "" then fSql = fSql & " AND xStatus = '" & request("xStatus") & "' "
		if request("memberId") <> "" then 
			fSql = fSql & " AND ( memberId LIKE '%" & request("memberId") & "%' OR realname LIKE '" & request("memberId") & "' "
			fSql = fSql & " OR nickname LIKE '%" & request("memberId") & "%') "
		end if
		fSql = fSql & " ORDER BY CuDTGeneric.xPostDate DESC "
		
		session("baseSql") = fSql
		session("cSql") = cSql
		
		
  elseif request("keep")="" then
   	
		cSql = "SELECT COUNT(*) FROM Pedia INNER JOIN CuDtGeneric ON CuDtGeneric.icuitem = Pedia.gicuitem "
		cSql = cSql & "INNER JOIN Member ON Pedia.memberId = Member.account "
		cSql = cSql & "WHERE CuDtGeneric.ictunit = " & iCtUnitId 
		
		fSql = "CuDTGeneric.*, Pedia.*, CodeMain_1.mValue AS isPublicValue, CodeMain.mValue AS isOpenValue, CodeMain_2.mValue AS cataValue, "
		fSql = fSql & "Member.realname, Member.nickname "
		fSql = fSql & "FROM CuDTGeneric LEFT OUTER JOIN CodeMain ON CuDTGeneric.vGroup = CodeMain.mCode "
		fSql = fSql & "AND CodeMain.codeMetaID = 'boolYN' LEFT OUTER JOIN CodeMain AS CodeMain_2 ON CuDTGeneric.topCat = CodeMain_2.mCode "
		fSql = fSql & "AND CodeMain_2.codeMetaID = 'pediacata' LEFT OUTER JOIN CodeMain AS CodeMain_1 "
		fSql = fSql & "ON CuDTGeneric.fCTUPublic = CodeMain_1.mCode AND CodeMain_1.codeMetaID = 'isPublic' "
		fSql = fSql & "INNER JOIN Pedia ON CuDTGeneric.icuitem = Pedia.gicuitem "
		fSql = fSql & "LEFT JOIN Member ON Pedia.memberId = Member.account "
		fSql = fSql & "WHERE CuDTGeneric.iCTUnit = " & iCtUnitId
		fSql = fSql & " ORDER BY CuDTGeneric.xPostDate DESC "
		
		session("baseSql") = fSql
		session("cSql") = cSql
	
	end if

	fSql = session("baseSql")	
	cSql = session("cSql")
	
	nowPage = Request.QueryString("nowPage")  '現在頁數
  PerPageSize = cint(Request.QueryString("pagesize"))
  if PerPageSize <= 0 then  PerPageSize = 50  

  set RSc = conn.execute(cSql)
  totRec = RSc(0)       '總筆數
  totPage = int(totRec / PerPageSize + 0.999)

  if cint(nowPage) < 1 then 
    nowPage = 1
  elseif cint(nowPage) > totPage then 
    nowPage = totPage 
  end if            	

	fSql = "SELECT TOP " & nowPage * PerPageSize & fSql

	Set RSreg = Server.CreateObject("ADODB.RecordSet")
	RSreg.CursorLocation = 2 ' adUseServer CursorLocationEnum
'	RSreg.CacheSize = PerPageSize

'----------HyWeb GIP DB CONNECTION PATCH----------
'	RSreg.Open fSql, Conn, 3, 1
Set RSreg =  Conn.execute(fSql)
	RSreg.CacheSize = PerPageSize
'----------HyWeb GIP DB CONNECTION PATCH----------


	if Not RSreg.eof then
		if totRec > 0 then 
      RSreg.PageSize = PerPageSize       '每頁筆數
      RSreg.AbsolutePage = nowPage      
		end if    
	end if   
%>

<body>
<div id="FuncName">
	<h1>資料管理／知識小百科</h1><font size=2>【目錄樹節點: 知識小百科】</font>
	<div id="Nav">
		<A href="PediaQuery.asp?ictunit=<%=iCtUnitId%>" title="指定查詢條件">查詢</A>		
		<A href="PediaAdd.asp?phase=add" title="新增資料">新增</A>		
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">	
	<font size=2>【主題單元:知識小百科】</font>
</div>
<Form id="Form2" name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
	<INPUT TYPE=hidden name=submitTask value="">	
	<div id="Page">
    <SPAN id=RunJob style="visibility:hidden"><input type=button class=cbutton value="刪除" id=button1 name=button1></SPAN>			
    <% if cint(nowPage) <> 1 then %>
			<img src="/images/arrow_previous.gif" alt="上一頁">       		
			<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> ，
    <% end if %>      
		共<em><%=totRec%></em>筆資料，每頁顯示
    <select id=PerPage size="1" style="color:#FF0000" class="select">            
			<option value="15"<%if PerPageSize=15 then%> selected<%end if%>>15</option>                       
      <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
      <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
      <option value="300"<%if PerPageSize=300 then%> selected<%end if%>>300</option>
    </select>     
		筆，目前在第
    <select id=GoPage size="1" style="color:#FF0000" class="select">
		<% For iPage=1 to totPage %> 
			<option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
    <% Next %>   
    </select>      
		頁	
    <% if cint(nowPage)<>totPage then %> 
     ，<a href="<%=HTProgPrefix%>List.asp?keep=Y&nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁
      <img src="/images/arrow_next.gif" alt="下一頁"></a> 
    <% end if %>     
	</div>

	<table cellspacing="0" id="ListTable">
	<tr>
		<th class="First" scope="col">來源連結</th>		
		<th><p align="center">標題(詞目)</p></th>
		<th><p align="center">分類</th>
		<th><p align="center">是否開放補充</th>
		<th><p align="center">發佈日期</th>
		<th><p align="center">會員名稱</th>
		<th><p align="center">審核狀態</th>
		<th><p align="center">狀態</th>
	</tr>
<%
	If not RSreg.eof then   
		for i = 1 to PerPageSize
			if RSreg("path") <> "" then
				xUrl = session("myWWWSiteURL") & RSreg("path")
			else
				xUrl = ""
			end if
			'xUrl = "http://www.boaf.gov.tw/boafwww/content.asp?mp=" & session("pvXdmp") & "&CuItem="  & RSreg("icuitem")
			pKey = "icuitem=" & RSreg("icuitem")
%>                  
			<tr>                  
				<td class=eTableContent>
					<% if xUrl <> "" Then %>
						<A href="<%=xUrl%>" target="_wMof">View</A></td>
					<% else %>
						&nbsp;</td>
					<% end if %>
				<td class=eTableContent><A href="/Pedia/PediaEdit.asp?<%=pKey%>&phase=edit&nowPage=<%=nowPage%>&pagesize=<%=PerPageSize%>"><%=RSreg("sTitle")%></A></td>
				<td class=eTableContent><%=RSreg("cataValue")%></td>
				<td class=eTableContent><%=RSreg("isOpenValue")%></td>
				<td class=eTableContent><%=RSreg("xPostDate")%></td>
				<td class=eTableContent>
				<%
						response.write RSreg("memberId") & "&nbsp; | &nbsp;" & RSreg("realname") & "&nbsp; | &nbsp;" & RSreg("nickname")
				%>
				</td>
				<td class=eTableContent>
				<%
						if RSreg("isPublicValue") = "公開" then
							response.write "通過"
						else
							response.write "不通過"
						end if
				%>
				</td>
				<td class=eTableContent><%=RSreg("xStatus")%></td>
			</tr>
<%
      RSreg.moveNext
      if RSreg.eof then exit for 
		next 
%>
	</table>       
  </form>  
</body>
</html>                                 
<% else %>
  <script language=vbs>
		msgbox "找不到資料, 請重設查詢條件!"
		'window.history.back
  </script>
<% end if %>

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
