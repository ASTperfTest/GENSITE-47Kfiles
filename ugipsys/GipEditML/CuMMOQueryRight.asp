<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="多媒體物件選取"
HTProgFunc="查詢"
HTUploadPath=session("Public")+"MMO"
session("MMOPath")=session("Public")+"MMO"
HTProgCode="GC1AP1"
HTProgPrefix="CuMMO" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

'----預設顯示圖示陣列
dim IconArray(3,1)
IconArray(0,0)="xxx":IconArray(0,1)="_audio"
IconArray(1,0)="wav":IconArray(1,1)="_flash"
IconArray(2,0)="zzz":IconArray(2,1)="_midi"
IconArray(3,0)="yyy":IconArray(3,1)="_video"

	xMMOitemCount = 0
		set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
		htPageDom.async = false
		htPageDom.setProperty("ServerHTTPRequest") = true	
	
		LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD") & "\sysPara.xml"
		xv = htPageDom.load(LoadXML)
  		if htPageDom.parseError.reason <> "" then 
	    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    			Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    			Response.End()
  		end if
		set MMONode=htPageDom.selectSingleNode("SystemParameter/MMOCtUnit")
		xMMO=nullText(MMONode)
		
if request("submitTask")="LIST" then
	fSql = "SELECT htx.icuitem,htx.stitle,htx.createdDate,C9.giCuItem,C9.mmofileName,C9.mmofolderName,C9.mmofileIcon,htx.ieditor,htx.deditDate,xref1.mvalue " & _
		"FROM CuDtx"+xMMO+" C9 " & _
		"Left Join CuDtGeneric htx ON htx.icuitem=C9.giCuItem " & _
		"LEFT JOIN CodeMain AS xref1 ON xref1.mcode = C9.mmofileType AND xref1.codeMetaId='fileType' " & _
		"Left Join Mmofolder MMO ON C9.mmofolderName=MMO.mmofolderName " & _
		"WHERE C9.mmofileName is not null and MMO.mmofolderName is not null "
	if request("htx_xKeyword")<>"" then
		fSql=fSql&" AND xkeyword Like '%"+request("htx_xKeyword")+"%'"
		xKeyword=request("htx_xKeyword")
	end if
	if request("htx_MMOFileName")<>"" then
		fSql=fSql&" AND mmofileName Like '%"+request("htx_MMOFileName")+"%'"
		xMMOFileName=request("htx_MMOFileName")
	end if
	if request("htx_MMOFolderName")<>"" then
		xMMOSiteID=left(request("htx_MMOFolderName"),instr(request("htx_MMOFolderName"),"---")-1)
		xMMOFolderName=mid(request("htx_MMOFolderName"),instr(request("htx_MMOFolderName"),"---")+3)
		fSql=fSql&" AND C9.mmofolderName = '"+xMMOFolderName+"'"
	end if	
	if request("htx_fileType")<>"" then
		fSql=fSql&" AND mmofileType Like '"+request("htx_fileType")+"%'"
		xfileType=request("htx_fileType")
	end if	
	if request("htx_IDateS") <> "" then
			rangeS = request("htx_IDateS")
			rangeE = request("htx_IDateE")
			if rangeE = "" then	rangeE=rangeS
		fSql = fSql & " AND htx.createdDate BETWEEN '"+rangeS+"' and '"+rangeE+"'"
	end if
'response.write fSql
'response.end	
	set RSreg=conn.execute(fSql)
end if
'response.write "xx="+session("mySiteURL")+session("MMOPath")+RSreg("MMOFolderName")+"<br>"
'response.write "yy="+HTUploadPath+RSreg("MMOFolderName")+"<br>"
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
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<table cellspacing="0">
  <tr>
    <td class="Label">物件檔案類型</td>
    <td>
<Select name="htx_fileType" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId='fileType' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
      </td>
    <td class="Label">物件存放目錄</td>
    <td>
<Select name="htx_MMOFolderName" size=1>
<option value="">全部</option>
			<%SQL="Select mmofolderId,mmofolderName,mmofolderDesc,mmositeId from Mmofolder order by mmofolderName"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF
				xMMOFolderDesc=""
				if not isNull(RSS("MMOFolderDesc")) and RSS("MMOFolderDesc")<>"" then xMMOFolderDesc="("&RSS("MMOFolderDesc")&")"
			%>

			<option value="<%=RSS("MMOSiteID")%>---<%=RSS(1)%>"><%=RSS(1)%><%=xMMOFolderDesc%></option>
			<%	RSS.movenext
			wend%>
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
    <td class="Label">建立日期</td>
    <td colspan=3>
   
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
	    	if j>=4 then
	    		response.write "</tr><tr>"
	    		j=0
	    	end if
  	    	xFileType=lcase(mid(RSreg("MMOFileName"), instrRev(RSreg("MMOFileName"), ".")+1))	        	    
%>
     	<td class=center height=100 width=60 onclick="editCall '<%=RSreg("iCUItem")%>','<%=RSreg("MMOFolderName")%>'">
     	  <table align=center>
     	  <%if not isNull(RSreg("MMOFileName")) then%>
     	    <tr><td align=center>
     	    <%if not isnull(RSreg("MMOFileIcon")) then%>
     	    	<img id="logo_MMOFileName" src="<%=session("mySiteURL")%>/public/MMO/<%=RSreg("MMOFolderName")%>/<%=RSreg("MMOFileIcon")%>">
     	    <%else
     	    	xFileExt="":xIcon=""
     	    	xPos=instrrev(RSreg("MMOFileName"),".")
     	    	if xPos<>0 then
     	    	    xFileExt=mid(RSreg("MMOFileName"),xPos+1)
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
     	    <%if RSreg("MMOFolderName")="/" then%>
     	    	<br><input type="radio" value="<%=HTUploadPath%><%=RSreg("MMOFolderName")%><%=RSreg("MMOFileName")%>|||<%=RSreg("sTitle")%>" name="giCuItem">
     	    <%else%>
     	    	<br><input type="radio" value="<%=HTUploadPath%><%=RSreg("MMOFolderName")%>/<%=RSreg("MMOFileName")%>|||<%=RSreg("sTitle")%>" name="giCuItem">
     	    <%end if%>
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
	<th scope="col">建立日期</td>
  </tr>
<%
    	if not RSreg.EOF then
    	    while not RSreg.EOF
				xMMOitemCount = xMMOitemCount + 1
%>
  <tr>
    <td class="Center">  
	<%if RSreg("MMOFolderName")="/" then%>
    		<input type="radio" value="<%=HTUploadPath%><%=RSreg("MMOFolderName")%><%=RSreg("MMOFileName")%>|||<%=RSreg("sTitle")%>" name="giCuItem">
     	<%else%>
    		<input type="radio" value="<%=HTUploadPath%><%=RSreg("MMOFolderName")%>/<%=RSreg("MMOFileName")%>|||<%=RSreg("sTitle")%>" name="giCuItem">
     	<%end if%>
    </td>
    <td><%=RSreg("sTitle")%></td>
    <td><%=RSreg("MMOFolderName")%></td>
    <td class="Center"><%=RSreg("MMOFileName")%></td>
    <td class="Center"><%=RSreg("mValue")%></td>
    <td class="Center"><%=RSreg("createdDate")%></td>
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
'  	if radiochk then 
'  		MMOPath = Left(form1.giCuItem(i).value,instr(form1.giCuItem(i).value,"|||")-1)
'  		xTitle = mid(form1.giCuItem(i).value,instr(form1.giCuItem(i).value,"|||")+3)
  	if not isNull(pickRadio) then 
  		MMOPath = Left(pickRadio.value,instr(pickRadio.value,"|||")-1)
  		xTitle = mid(pickRadio.value,instr(pickRadio.value,"|||")+3)
		window.opener.doInsertImageRight MMOPath,xTitle
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
    	PopForm.htx_fileType.value="<%=xFileType%>"
    	PopForm.htx_MMOFolderName.value="<%=request("htx_MMOFolderName")%>"
    	PopForm.htx_MMOFileName.value="<%=xMMOFileName%>"
    	PopForm.htx_xKeyword.value="<%=xKeyword%>"
    	PopForm.htx_IDateS.value="<%=rangeS%>"
    	PopForm.pcShowhtx_IDateS.value="<%=d7date(rangeS)%>"
    	PopForm.htx_IDateE.value="<%=rangeE%>"
    	PopForm.pcShowhtx_IDateE.value="<%=d7date(rangeE)%>"
    <%end if%>
end sub

sub window_onLoad  
	clientInitForm
end sub 
   
sub formSearchSubmit()
'alert PopForm.htx_IDateS.value+"[]"+PopForm.htx_IDateE.value
         PopForm.submitTask.value="LIST"
         PopForm.listWay.value=""
         PopForm.Submit
end Sub

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

