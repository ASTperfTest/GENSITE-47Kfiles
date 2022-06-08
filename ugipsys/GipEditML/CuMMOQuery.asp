<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="多媒體物件選取"
HTProgFunc="查詢"
HTProgCode="GC1AP1"
HTProgPrefix="CuMMO" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function MMOPathStr(MMOFolderID)	'----940407 MMO路徑字串
	sql = "SELECT MM.mmositeId,mmofolderParent,Case mmofolderParent when 'zzz' then mmositeName else mmofolderNameShow END MMOFolderNameShow " & _ 
		"FROM Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId where mmofolderID=" & MMOFolderID
	set RSN = conn.execute(sql)
	xParent = RSN("mmofolderParent")
	xPathStr = RSN("MMOFolderNameShow")
	xMMOSiteID = RSN("mmositeId")
	while xParent <> "zzz"
		sql = "SELECT MM.mmositeId,mmofolderParent,Case mmofolderParent when 'zzz' then mmositeName else mmofolderNameShow END MMOFolderNameShow " & _ 
			"FROM Mmofolder MM Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId where MM.mmositeId=" & pkstr(xMMOSiteID,"") & " and mmofolderName=" & pkstr(xParent,"")
		set RS = conn.execute(sql)
		xPathStr = RS("MMOFolderNameShow") & " / " & xPathStr
		xParent = RS("mmofolderParent")
		xMMOSiteID = RSN("mmositeId")
	wend
	MMOPathStr = xPathStr
end function


'----預設顯示圖示陣列
dim IconArray(3,1)
IconArray(0,0)="xxx":IconArray(0,1)="_audio"
IconArray(1,0)="wav":IconArray(1,1)="_flash"
IconArray(2,0)="zzz":IconArray(2,1)="_midi"
IconArray(3,0)="yyy":IconArray(3,1)="_video"

xMMOitemCount = 0
imgPos=request("imgpos")
	
if request("submitTask")="LIST" then
	SQLM="Select mmositeId+mmofolderName as MMOFOlderID from Mmofolder where mmofolderID=" & request("htx_MMOFolderID")
	Set RSM=conn.execute(SQLM)
	SQLM2="Select sbaseTableName " & _
		"from CtUnit C Left Join BaseDsd B ON C.ibaseDsd=B.ibaseDsd " & _
		"where C.CtUnitId=" & request("htx_CtUnitID")
	Set RSM2=conn.execute(SQLM2)
	if not RSM.eof then xMMOFolderID=RSM("MMOFOlderID")
	fSql = "SELECT htx.icuitem,htx.stitle,htx.createdDate,C9.*, " & _
		"htx.ieditor,htx.deditDate,xref1.mvalue,xref1.htmlTag,MM.mmofolderName,MM.mmositeId,MS.upLoadSiteHttp " & _
		"FROM "&RSM2("sbaseTableName")&" C9 " & _
		"Left Join CuDtGeneric htx ON htx.icuitem=C9.gicuitem " & _
		"LEFT JOIN MMOFileType AS xref1 ON xref1.mcode = C9.mmoFileType " & _
		"Left Join Mmofolder MM ON C9.mmofolderId=MM.mmofolderId " & _
		"Left Join Mmosite MS ON MM.mmositeId=MS.mmositeId " & _
		"Left Join CtUnit C ON htx.ictunit=C.CtUnitId " & _
		"WHERE C9.mmofileName is not null and MM.mmofolderName is not null "
	if request("htx_xKeyword")<>"" then
		fSql=fSql&" AND xkeyword Like '%"+request("htx_xKeyword")+"%'"
		xKeyword=request("htx_xKeyword")
	end if
	if request("htx_MMOFileName")<>"" then
		fSql=fSql&" AND C9.mmofileName Like '%"+request("htx_MMOFileName")+"%'"
		xMMOFileName=request("htx_MMOFileName")
	end if
	if request("htx_CtUnitID")<>"" then
		fSql=fSql&" AND ictunit = "+request("htx_CtUnitID")
		xCtUnitID=request("htx_CtUnitID")
	end if	
	if request("htx_MMOFolderID")<>"" then
		fSql=fSql&" AND CASE mmofolderName WHEN 'zzz' THEN mmofolderName ELSE MM.mmositeId + mmofolderName END like '"&xMMOFolderID&"%' "
		xMMOFolderID=request("htx_MMOFolderID")
	end if	
	if request("htx_fileType")<>"" then
		fSql=fSql&" AND mmoFileType Like '"+request("htx_fileType")+"%'"
		xfileType=request("htx_fileType")
	end if	
	if request("htx_IDateS") <> "" then
			rangeS = request("htx_IDateS")
			rangeE = request("htx_IDateE")
			if rangeE = "" then	rangeE=rangeS
		fSql = fSql & " AND htx.createdDate BETWEEN '"+rangeS+"' and '"+rangeE+"'"
	end if
	fSql=fSql&" Order by ictunit,MM.mmofolderId"
'response.write fSql
'response.end	
	set RSreg=conn.execute(fSql)
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>插入圖檔</title>
<link href="css/popup.css" rel="stylesheet" type="text/css">
<link href="css/editor.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="PopFormName">多媒體物件查詢 - 圖檔查詢</div>
<form action="" method="" id="PopForm">
<INPUT TYPE=hidden name=submitTask value="">
<INPUT TYPE=hidden name=listWay value="">
<INPUT TYPE=hidden name=imgPos value="<%=request.querystring("imgpos")%>">
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<table cellspacing="0">
  <tr>
    <td class="Label">主題單元</td>
    <td>
	<Select name="htx_CtUnitID" size=1 onchange="VBS:MMOFolderIDSelect()">
	<option value="">請選擇</option>
		<%SQL="Select ctUnitId,ctUnitName from CtUnit CT Left Join baseDsd B ON Ct.ibaseDsd=B.ibaseDsd where rdsdcat='MMO'"
		SET RSS=conn.execute(SQL)
		While not RSS.EOF%>

		<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
		<%	RSS.movenext
		wend%>
	</select>
      </td>  
    <td class="Label">物件存放目錄</td>
    <td>
	<Select name="htx_MMOFolderID" size=1>
	</select>
    </td>
  </tr>
  <tr>
    <td class="Label">物件檔案名稱</td>
    <td><input name="htx_MMOFileName" type="text" class="InputText" value="" size="20"></td>
    <td class="Label">關鍵詞</td>
    <td><input name="htx_xKeyword" type="text" class="InputText" value="" size="20"></td>
  </tr>
  <tr>
    <td class="Label">物件檔案類型</td>
    <td>
<Select name="htx_fileType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from MMOFileType Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
      </td>  
    <td class="Label">建立/編修日期</td>
    <td>   
    <input name="pcShowhtx_IDateS" type="text" class="InputText" size="10" readonly onclick="VBS: popCalendar 'htx_IDateS','htx_IDateE'">
      ～<input name="htx_IDateS" type=hidden><input name="htx_IDateE" type=hidden>
        <input name="pcShowhtx_IDateE" type="text" class="InputText" size="10" readonly onclick="VBS: popCalendar 'htx_IDateE',''">
        </td>
  </tr>
</table>
<hr>
<input type="button" class="InputButton" value="查詢" onClick="formSearchSubmit()">
<input type="button" class="InputButton" value="重填" onClick="resetForm()">
</form>
<%if request("submitTask")="LIST" then%>
<form name="form1" method="" action="">
<table width="100%" cellspacing="0" id="PopList">
  <tr>
	<%if request("listWay")="icon" then%>
  	<td colspan=6 class=right>
  	<input type="button" class="InputButton" value="清單顯示" onClick="listForm()">
  	<input type="button" class="InputButton" value="新增物件" onClick="AddMMO()">
  	</td>
	<%else%>
  	<td colspan=6 class=right>
  	<input type="button" class="InputButton" value="圖示顯示" onClick="iconForm()">
  	<input type="button" class="InputButton" value="新增物件" onClick="AddMMO()">
  	</td>
	<%end if%>  
  </tr>
    <%	
    if request("listWay")="icon" then
    	if not RSreg.EOF then%>
  <tr>
    		<% 
    	    while not RSreg.EOF
		xMMOitemCount = xMMOitemCount + 1
		if not isnull(RSreg("upLoadSiteHttp")) then
			httpStr=RSreg("upLoadSiteHttp")
			if right(httpStr,1)<>"/" then httpStr = httpStr & "/"	
			httpStr=httpStr+RSreg("mmositeId")+RSreg("mmofolderName")+"/"+RSreg("mmoFileName")
		else
			httpStr=session("MMOPublic")+RSreg("mmositeId")+RSreg("mmofolderName")+"/"+RSreg("mmoFileName")
		end if
		if not isNull(RSreg("htmlTag")) then		
			xHTMLTag=RSreg("htmlTag")
			xHTMLTag=replace(xHTMLTag,"{MMOPATHSTR}","src='"&httpStr&"'")
			xHTMLTag=replace(xHTMLTag,"{MMOALTSTR}","alt='"&RSreg("stitle")&"'")
			if not isNull(RSreg("mmoFileHeight")) then 
				xHTMLTag=replace(xHTMLTag,"{MMOHEIGHTSTR}","height='"&RSreg("mmoFileHeight")&"'")
			else
				xHTMLTag=replace(xHTMLTag,"{MMOHEIGHTSTR}","")		
			end if
			if not isNull(RSreg("mmoFileWidth")) then 
				xHTMLTag=replace(xHTMLTag,"{MMOWIDTHSTR}","width='"&RSreg("mmoFileWidth")&"'")
			else
				xHTMLTag=replace(xHTMLTag,"{MMOWIDTHSTR}","")
			end if
			xHTMLTag=replace(xHTMLTag,"{MMOIDSTR}","MMOID='"&RSreg("icuitem")&"'")		
		else
			xHTMLTag=""
		end if			
	    	if j>=5 then
	    		response.write "</tr><tr>"
	    		j=0
	    	end if
  	    	xFileType=lcase(mid(RSreg("mmoFileName"), instrRev(RSreg("mmoFileName"), ".")+1))	        	    
%>
     	<td class=center height=100 width=60>
     	  <table align=center>
     	  <%if not isNull(RSreg("mmoFileName")) then%>
     	    <tr><td align=center>
     	    <%if not isnull(RSreg("mmoFileIcon")) then%>
     	    	<img id="logo_MMOFileName" width=40 height=30 src="<%=session("MMOPublic")%><%=RSreg("mmositeId")%><%=RSreg("mmofolderName")%>/<%=RSreg("mmoFileIcon")%>">
     	    <%else
     	    	xFileExt="":xIcon=""
     	    	xPos=instrrev(RSreg("mmoFileName"),".")
     	    	if xPos<>0 then
     	    	    xFileExt=mid(RSreg("mmoFileName"),xPos+1)
     	    	    for k=0 to ubound(IconArray)
     	    	    	if xFileExt=IconArray(k,0) then
     	    	    	    xIcon=IconArray(k,1)
     	    	    	    exit for
     	    	    	end if
     	    	    next
     	    	end if
     	    %>
     	    	<img id="logo_MMOFileName" src="../PageStyle/thumb_media<%=xIcon%>.gif">    	    
     	    <%end if%>
     	    <input type="radio" value="<%=xHTMLTag%>" name="giCuItem"><font size=2><%=RSreg("mmoFileName")%></font>
     	    </td></tr> 
     	  <%else%>
     	    <tr><td class=center>&nbsp;</td></tr>
     	  <%end if%>
     	  </table>
     	</td>         
<%	    j=j+1    	    
	    	RSreg.movenext	    
	    wend%>
  </tr>
	<%end if	
    else%>  
  <tr>
	<th scope="col">&nbsp;</td>
	<th scope="col">標題</td>
	<th scope="col">存放目錄</td>
	<th scope="col">物件檔案名稱</td>
	<th scope="col">檔案類型</td>
	<th scope="col">建立/編修日期</td>
  </tr>
<%
    	if not RSreg.EOF then
    	    while not RSreg.EOF
		xMMOitemCount = xMMOitemCount + 1
		if not isnull(RSreg("upLoadSiteHttp")) then
			httpStr=RSreg("UpLoadSiteHttp")
			if right(httpStr,1)<>"/" then httpStr = httpStr & "/"	
			httpStr=httpStr+RSreg("mmositeId")+RSreg("mmofolderName")+"/"+RSreg("mmoFileName")
		else
			httpStr=session("MMOPublic")+RSreg("mmositeId")+RSreg("mmofolderName")+"/"+RSreg("mmoFileName")
		end if
		if not isNull(RSreg("htmlTag")) then				
			xHTMLTag=RSreg("htmlTag")
			xHTMLTag=replace(xHTMLTag,"{MMOPATHSTR}",httpStr)
			xHTMLTag=replace(xHTMLTag,"{MMOALTSTR}",RSreg("sTitle"))
			if not isNull(RSreg("mmoFileHeight")) then 
				xHTMLTag=replace(xHTMLTag,"{MMOHEIGHTSTR}","height='"&RSreg("mmoFileHeight")&"'")
			else
				xHTMLTag=replace(xHTMLTag,"{MMOHEIGHTSTR}","")		
			end if
			if not isNull(RSreg("mmoFileWidth")) then 
				xHTMLTag=replace(xHTMLTag,"{MMOWIDTHSTR}","width='"&RSreg("mmoFileWidth")&"'")
			else
				xHTMLTag=replace(xHTMLTag,"{MMOWIDTHSTR}","")
			end if
			xHTMLTag=replace(xHTMLTag,"{MMOIDSTR}",RSreg("icuitem"))
		else
			xHTMLTag=""
		end if	
%>
  <tr>
    <td class="Center">  
    		<input type="radio" value="<%=xHTMLTag%>" name="giCuItem">
    </td>
    <td><%=RSreg("stitle")%></td>
    <td><%=MMOPathStr(RSreg("mmofolderId"))%></td>
    <td class="Center"><%=RSreg("mmoFileName")%></td>
    <td class="Center"><%=RSreg("mvalue")%></td>
    <td class="Center"><%=RSreg("deditDate")%></td>
  </tr>
<%	     	RSreg.movenext
	    wend
    	end if
    end if
    %>
</table>
  <input type="button" class="InputButton" value="確定" onClick="closeForm()">
</form>
<%end if%>
</body>
</html>
<script language=vbs>
imgPos="<%=imgPos%>"
sub AddMMO()
	window.navigate "CuMMOAdd.asp?phase=add","","scrollbars=yes,width=700,height=500"
end sub

sub closeForm()
		xMMOitemCount = <%=xMMOitemCount%>
        pickRadio = null
  if xMMOitemCount>1 then
  	for i=0 to form1.giCuItem.length-1                           
   	    if form1.giCuItem(i).checked then 
     		set pickRadio =  form1.giCuItem(i)                       
      		exit for                           
   	    end if                           
  	next  
  else
  		if form1.giCuItem.checked then	
     		set pickRadio =  form1.giCuItem                      
  		end if
  end if
  	if not isNull(pickRadio) then 
  		if instr(pickRadio.value,"<IMG")<>0 then
  		    if imgPos="left" then
  			MMOStr=replace(pickRadio.value," MMOID"," style='float:left' MMOID") 
  		    elseif imgPos="right" then
   			MMOStr=replace(pickRadio.value," MMOID"," style='float:right' MMOID")  	
   		    else	    
  			MMOStr=pickRadio.value
  		    end if
  		else
  			MMOStr=pickRadio.value
  		end if
		window.opener.doInsertImage MMOStr
		window.close
	else
		alert "請選取物件!"
		exit sub
	end if
end sub

sub resetForm
	PopForm.reset()
end sub

sub clientInitForm
    <%if request("submitTask")="LIST" then%>
    	PopForm.htx_CtUnitID.value="<%=xCtUnitID%>"
    	PopForm.htx_fileType.value="<%=xFileType%>"
    	PopForm.htx_MMOFileName.value="<%=xMMOFileName%>"
    	PopForm.htx_xKeyword.value="<%=xKeyword%>"
    	PopForm.htx_IDateS.value="<%=rangeS%>"
    	PopForm.pcShowhtx_IDateS.value="<%=d7date(rangeS)%>"
    	PopForm.htx_IDateE.value="<%=rangeE%>"
    	PopForm.pcShowhtx_IDateE.value="<%=d7date(rangeE)%>"
    	MMOFolderIDSelect()
    	PopForm.htx_MMOFolderID.value="<%=xMMOFolderID%>"
    <%end if%>
end sub

sub window_onLoad  
	clientInitForm
end sub 
   
sub formSearchSubmit()
	if PopForm.htx_CtUnitID.value="" then
		alert "請選擇主題單元!"
		PopForm.htx_CtUnitID.focus
		exit sub
	end if
         PopForm.submitTask.value="LIST"
         PopForm.listWay.value=""
         PopForm.Submit
end Sub

sub MMOFolderIDSelect()
 	set xsrc = document.all("htx_MMOFolderID")
 	removeOption(xsrc)
 	set oXML = createObject("Microsoft.XMLDOM")
 	oXML.async = false
 	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_MMOFolderID.asp?CtUnitID=" & PopForm.htx_CtUnitID.value
 	oXML.load(xURI)
 	set pckItemList = oXML.selectNodes("divList/row")
 	for each pckItem in pckItemList
  		xaddOption xsrc, pckItem.selectSingleNode("mValue").text,pckItem.selectSingleNode("mCode").text
 	next
 	xsrc.selectedIndex = 0
end sub	
sub xaddOption(xlist,name,value)
 	set xOption = document.createElement("OPTION")
 	xOption.text=name
 	xOption.value=value
 	xlist.add(xOption)
end sub
sub removeOption(xlist)
 	for i=xlist.options.length-1 to 0 step -1
  		xlist.options.remove(i)
 	next
 	xlist.selectedIndex = -1
end sub

sub listForm()
         PopForm.submitTask.value="LIST"
         PopForm.listWay.value=""
         PopForm.Submit
end Sub

sub iconForm()
         PopForm.submitTask.value="LIST"
         PopForm.listWay.value="icon"
         PopForm.Submit
end Sub

Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=d7date(document.all.CalendarTarget.value)
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=d7date(document.all.CalendarTarget.value)
        end if
	end if
end sub   
</script>

