<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="目錄樹管理"
HTProgFunc="編修節點"
HTUploadPath="/public/"
HTProgCode="GE1T21"
HTProgPrefix="CtUnitNode" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<html>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim pKey
Dim xchildCount
Dim RSreg
Dim formFunction
taskLable="編輯" & HTProgCap

goFlag = checkGIPconfig("RSSsource")

sub oldFrank
goFlag = false
fSql = "sp_columns 'CuDTGeneric'"
set RSlist = conn.execute(fSql)
if not RSlist.EOF then
    while not RSlist.eof
	if RSlist("COLUMN_NAME") = "KMautoID" then
		goFlag = true
	end if
      	RSlist.moveNext
    wend
end if
end sub

 apath=server.mappath(HTUploadPath) & "\"
if request.querystring("phase")<>"edit" then
Set xup = Server.CreateObject("TABS.Upload")
xup.codepage=65001
xup.Start apath

else
Set xup = Server.CreateObject("TABS.Upload")
end if

function xUpForm(xvar)
	xUpForm = xup.form(xvar)
end function

if xUpForm("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
	end if

elseif xUpForm("submitTask") = "DELETE" then
	SQL = "DELETE FROM CatTreeNode WHERE ctNodeId=" & pkStr(request.queryString("ctNodeId"),"")
	conn.Execute SQL
	'======	2006.5.9 by Gary
	if checkGIPconfig("RSSandQuery") then  
		SQL = "DELETE FROM RSSPool WHERE iCtNode=" & pkStr(request.queryString("ctNodeId"),"")
		conn.Execute SQL
	end if
	'======	
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlCom = "SELECT htx.*, b.sbaseDsdname AS refibaseDsd, r1.mvalue AS refctUnitKind, r2.mvalue AS refctNodeKind " _
		& " ,(Select count(*) from CatTreeNode where dataParent=htx.ctNodeId) childCount" _
		& " FROM CatTreeNode AS htx LEFT JOIN CtUnit AS u ON u.ctUnitId=htx.ctUnitId" _
		& " LEFT JOIN BaseDsd AS b ON b.ibaseDsd=u.ibaseDsd" _
		& " LEFT JOIN CodeMainLong AS r1 ON r1.mcode=u.ctUnitKind AND r1.codeMetaId=N'refCTUKind'" _
		& " LEFT JOIN CodeMain AS r2 ON r2.mcode=htx.ctNodeKind AND r2.codeMetaId='refctNodeKind'" _
		& " WHERE htx.ctNodeId=" & pkStr(request.queryString("ctNodeId"),"")
	Set RSreg = Conn.execute(sqlcom)
	'======	2006.5.9 by Gary
	if checkGIPconfig("RSSandQuery") then 
		session("YNrss") = RSreg("YNrss")
	end if
	'======	
	xchildCount = RSreg("childCount")
	pKey = ""
	pKey = pKey & "&ctNodeId=" & RSreg("ctNodeId")
	if pKey<>"" then  pKey = mid(pKey,2)

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
'	on error resume next
	if xUpForm("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if xUpForm("htx_"&fldName) <> "" then
			xValue = xUpForm("htx_"&fldName)
		end if
	end if
'	if err.number > 0 then	
'		response.write "***"&fldName&"***"&Err.Description&"***"
'		err.number = 0
'		Err.Clear
'	end if 
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<%Sub initForm() %>
    <script language=vbs>
      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

     sub clientInitForm
	reg.htx_ctNodeKind.value= "<%=qqRS("ctNodeKind")%>"
	reg.htx_ctRootId.value= "<%=qqRS("ctRootId")%>"
	reg.htx_editUserId.value= "<%=session("userID")%>"
	reg.htx_editDate.value= "<%=date()%>"
	reg.htx_ctNodeId.value= "<%=qqRS("ctNodeId")%>"
	reg.htx_catName.value= "<%=qqRS("catName")%>"
	reg.htx_ctUnitId.value= "<%=qqRS("ctUnitId")%>"
	reg.dsp_refctUnitKind.value= "<%=qqRS("refctUnitKind")%>"
	reg.dsp_refibaseDsd.value= "<%=qqRS("refibaseDsd")%>"
	initImgFile "CtNameLogo","<%=qqRS("CtNameLogo")%>"
	reg.htx_catShowOrder.value= "<%=qqRS("catShowOrder")%>"
	reg.htx_inUse.value= "<%=qqRS("inUse")%>"
	reg.htx_xslList.value= "<%=qqRS("xslList")%>"
	reg.htx_xslData.value= "<%=qqRS("xslData")%>"
	reg.htx_dcondition.value= "<%=qqRS("dcondition")%>"
	<%if RSreg("ctNodeKind")="C" then%>
		reg.htx_ctNodeNpkind.value= "<%=qqRS("ctNodeNpkind")%>"
	<%end if%>
	<%if goFlag then%>
		reg.htx_RSSURLID.value= "<%=qqRS("RSSURLID")%>"
		reg.htx_RSSNodeType.value= "<%=qqRS("RSSNodeType")%>"
	<%end if%>		
	<% if checkGIPconfig("KMCat") then %>
		reg.htx_KMautoID.value= "<%=qqRS("KMautoID")%>"
		reg.htx_KMCatID.value= "<%=qqRS("KMCatID")%>"
		reg.htx_KMCat.value= "<%=qqRS("KMCat")%>"
	<%end if%>		
	<% if checkGIPconfig("SubjectCat") then %>
		reg.htx_SubjectID.value= "<%=qqRS("SubjectID")%>"
	<%end if%>
	<%'======	2006.3.30 by Gary
	 if checkGIPconfig("RSSandQuery") then  %>
		reg.htx_YNrss.value= "<%=qqRS("YNrss")%>"
		reg.htx_YNquery.value= "<%=qqRS("YNquery")%>"
		
	<%end if
		'====== %>					
	<%if checkGIPconfig("SubjectMonth") then%>	
		initCheckbox "bfx_SubjectMonth","<%=qqRS("SubjectMonth")%>"
	<%end if%>				
	
    end sub

    sub window_onLoad
         clientInitForm
    end sub
    
    sub initRadio(xname,value)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    end sub

    sub initOtherRadio(xname,value, otherName)
    	for each xri in reg.all(xname)
    		if xri.value=value then 
    			xri.checked = true
    			exit sub
    		end if
    	next
    	if value="" then	exit sub
		reg.all(xname).item(reg.all(xname).length-1).checked = true
		reg.all(otherName).value = value
		reg.all(xname).item(reg.all(xname).length-1).value = value
    end sub

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
    
    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,", ")
    	valueCount = ubound(valueArray) + 1
    	value = ckValue & ","
    	ckCount = 0
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    			ckCount = ckCount + 1
    		end if
    	next
		if ckCount <> valueCount then
			reg.all(xname).item(reg.all(xname).length-1).checked = true
			reg.all(otherName).value = valueArray(ubound(valueArray))
			reg.all(xname).item(reg.all(xname).length-1).value = valueArray(ubound(valueArray))
		end if
    end sub

    sub initImgFile(xname, value)
		reg.all("htImgActCK_"&xname).value=""
		reg.all("htImg_"&xname).style.display="none"
		reg.all("hoImg_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).src= "<%=HTUploadPath%>" & value
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addLogo(xname)	'新增logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="addLogo"
End sub

Sub chgLogo(xname)	'更換logo
    reg.all("htImg_"&xname).style.display=""
    reg.all("htImgActCK_"&xname).value="editLogo"
End sub

Sub orgLogo(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value=""
End sub

Sub delLogo(xname)	'刪除logo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htImg_"&xname).value=""
    reg.all("htImg_"&xname).style.display="none"
    reg.all("htImgActCK_"&xname).value="delLogo"
End sub

    sub initAttFile(xname, value, orgValue)
		reg.all("htFileActCK_"&xname).value=""
		reg.all("htFile_"&xname).style.display="none"
		reg.all("hoFile_"&xname).value = value
		document.all("LbtnHide2_"&xname).style.display="none"
		If value="" then
			document.all("noLogo_"&xname).style.display=""
	   		document.all("logo_"&xname).style.display="none"
	   		document.all("LbtnHide0_"&xname).style.display=""
	   		document.all("LbtnHide1_"&xname).style.display="none"
		Else
	   		document.all("logo_"&xname).style.display=""
	   		document.all("logo_"&xname).innerText= orgValue
	   		document.all("LbtnHide0_"&xname).style.display="none"
	   		document.all("LbtnHide1_"&xname).style.display=""
	   		document.all("noLogo_"&xname).style.display="none"
		End if
	end sub

Sub addXFile(xname)	'新增logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="addLogo"
End sub

Sub chgXFile(xname)	'更換logo
    reg.all("htFile_"&xname).style.display=""
    reg.all("htFileActCK_"&xname).value="editLogo"
End sub

Sub orgXFile(xname)	'原圖
    document.all("LbtnHide2_"&xname).style.display="none"
    document.all("LbtnHide1_"&xname).style.display=""
    document.all("logo_"&xname).style.display=""
    document.all("noLogo_"&xname).style.display="none"
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value=""
End sub

Sub delXFile(xname)	'刪除logo
    document.all("LbtnHide2_"&xname).style.display=""
    document.all("LbtnHide1_"&xname).style.display="none"
    document.all("logo_"&xname).style.display="none"
    document.all("noLogo_"&xname).style.display=""
    reg.all("htFile_"&xname).value=""
    reg.all("htFile_"&xname).style.display="none"
    reg.all("htFileActCK_"&xname).value="delLogo"
End sub

 </script>	
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
     <script language="vbscript">
      sub formModSubmit()
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
'     Exit Sub
'  End if
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
  
	IF reg.htx_ctNodeKind.value = Empty Then 
		MsgBox replace(nMsg,"{0}","節點類別"), 64, "Sorry!"
		reg.htx_ctNodeKind.focus
		exit sub
	END IF
	IF (reg.htx_ctRootId.value <> "") AND (NOT isNumeric(reg.htx_ctRootId.value)) Then
		MsgBox replace(iMsg,"{0}","目錄樹ID"), 64, "Sorry!"
		reg.htx_ctRootId.focus
		exit sub
	END IF
	IF (reg.htx_editDate.value <> "") AND (NOT isDate(reg.htx_editDate.value)) Then
		MsgBox replace(dMsg,"{0}","編輯日期"), 64, "Sorry!"
		reg.htx_editDate.focus
		exit sub
	END IF
	IF (reg.htx_ctNodeId.value <> "") AND (NOT isNumeric(reg.htx_ctNodeId.value)) Then
		MsgBox replace(iMsg,"{0}","目錄樹節點ID"), 64, "Sorry!"
		reg.htx_ctNodeId.focus
		exit sub
	END IF
	IF reg.htx_catName.value = Empty Then 
		MsgBox replace(nMsg,"{0}","標題"), 64, "Sorry!"
		reg.htx_catName.focus
		exit sub
	END IF
	IF blen(reg.htx_catName.value) > 100 Then
		MsgBox replace(replace(lMsg,"{0}","標題"),"{1}","30"), 64, "Sorry!"
		reg.htx_catName.focus
		exit sub
	END IF
	IF blen(reg.htx_dcondition.value) > 100 Then
		MsgBox replace(replace(lMsg,"{0}","限制條件"),"{1}","100"), 64, "Sorry!"
		reg.htx_dcondition.focus
		exit sub
	END IF
	IF (reg.htx_ctUnitId.value <> "") AND (NOT isNumeric(reg.htx_ctUnitId.value)) Then
		MsgBox replace(iMsg,"{0}","主題單元ID"), 64, "Sorry!"
		reg.htx_ctUnitId.focus
		exit sub
	END IF
	IF reg.htImg_CtNameLogo.value <> "" Then
		xIMGname = reg.htImg_CtNameLogo.value
		xFileType = ""
		if instr(xIMGname, ".")>0 then	xFileType=lcase(mid(xIMGname, instr(xIMGname, ".")+1))
		IF xFileType<>"gif" AND xFileType<>"jpg" AND xFileType<>"jpeg" then
			MsgBox replace(pMsg,"{0}","標題圖示"), 64, "Sorry!"
			reg.htImg_CtNameLogo.focus
			exit sub
		END IF
	END IF
	IF reg.htx_inUse.value = Empty Then 
		MsgBox replace(nMsg,"{0}","是否開放"), 64, "Sorry!"
		reg.htx_inUse.focus
		exit sub
	END IF
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("注意！"& vbcrlf & vbcrlf &"　你確定刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>


<!--#include file="CtUnitNodeFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
	    <a href="VBScript: history.back">回上一頁</a>
	    <!--input type=button class=cbutton value="產生DND" id=button1-->
		<%if (HTProgRight and 8)=8 then%>
			<A href="CtNodeDND.asp?ctNodeId=<%=RSreg("ctNodeId")%>">DND</A>
		<%end if%>		
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName"><%=HTProgFunc%></div>

<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li><span class="Must">*</span>為必要欄位</li>
	</ul>
</div>
</body>
</html>
<%End sub '--- showHTMLTail() ------%>


<%Sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
'  			window.history.back
	</script>
<%
    End sub '---- showErrBox() ----

Sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = '"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "「統一編號」重複!!請重新輸入客戶統一編號!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()
	sql = "UPDATE CatTreeNode SET "
		sql = sql & "ctNodeKind=" & pkStr(xUpForm("htx_ctNodeKind"),",")
		sql = sql & "ctRootId=" & pkStr(xUpForm("htx_ctRootId"),",")
		sql = sql & "editUserId=" & pkStr(xUpForm("htx_editUserId"),",")
		sql = sql & "editDate=" & pkStr(xUpForm("htx_editDate"),",")
		sql = sql & "catName=" & pkStr(xUpForm("htx_catName"),",")
		sql = sql & "ctUnitId=" & pkStr(xUpForm("htx_ctUnitId"),",")
		sql = sql & "xslList=" & pkStr(xUpForm("htx_xslList"),",")
		sql = sql & "xslData=" & pkStr(xUpForm("htx_xslData"),",")
		sql = sql & "dcondition=" & pkStr(xUpForm("htx_dcondition"),",")
		if goFlag then
			sql = sql & "RSSURLID=" & pkStr(xUpForm("htx_RSSURLID"),",")
			sql = sql & "RSSNodeType=" & pkStr(xUpForm("htx_RSSNodeType"),",")
	    end if
	    if checkGIPconfig("KMCat") then
	     	sql = sql & "KMautoID=" & pkStr(xUpForm("htx_KMautoID"),",")
			sql = sql & "KMCatID=" & pkStr(xUpForm("htx_KMCatID"),",")
			sql = sql & "KMCat=" & pkStr(xUpForm("htx_KMCat"),",")
		end if
		if checkGIPconfig("SubjectCat") then
			sql = sql & "SubjectID=" & pkStr(xUpForm("htx_SubjectID"),",")
		end if
	    if checkGIPconfig("SubjectMonth") then	
		sql = sql & "SubjectMonth=" & pkStr(xUpForm("bfx_SubjectMonth"),",")
	    end if	
	IF xUpForm("htImgActCK_CtNameLogo") <> "" Then
	  actCK = xUpForm("htImgActCK_CtNameLogo")
	  if actCK="editLogo" OR actCK="addLogo" then
		fname = ""
		For each xatt in xup.Attachments
		  if xatt.Name = "htImg_CtNameLogo" then
			ofname = xatt.FileName
			fnExt = ""
			if instr(ofname, ".")>0 then fnext=mid(ofname, instr(ofname, "."))
			tstr = now()
			nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
			sql = sql & "CtNameLogo=" & pkStr(nfname,",")
			IF xUpForm("hoImg_CtNameLogo") <> "" Then _
				xup.DeleteFile apath & xUpForm("hoImg_CtNameLogo")
			xatt.SaveFile apath & nfname, false
		  end if
		Next
	  elseif actCK="delLogo" then
		xup.DeleteFile apath & xUpForm("hoImg_CtNameLogo")
		sql = sql & "CtNameLogo=null,"
	  end if
	END IF
		sql = sql & "catShowOrder=" & pkStr(xUpForm("htx_catShowOrder"),",")
		sql = sql & "inUse=" & pkStr(xUpForm("htx_inUse"),",")
'======	2006.3.30 by Gary		
	IF xUpForm("htx_YNrss") <> "" Then
		sql = sql & "YNrss=" & pkStr(xUpForm("htx_YNrss"),",")
	END IF	
	IF xUpForm("htx_YNquery") <> "" Then
		sql = sql & "YNquery=" & pkStr(xUpForm("htx_YNquery"),",")
	END IF	
'======		
		sql = sql & "ctNodeNpkind=" & pkStr(xUpForm("htx_ctNodeNpkind"),",")
		
	sql = left(sql,len(sql)-1) & " WHERE ctNodeId=" & pkStr(request.queryString("ctNodeId"),"")
	conn.Execute(SQL)  

	'======	2006.5.9 by Gary
	if checkGIPconfig("RSSandQuery") then 
		IF xUpForm("htx_YNrss") = "N" Then
			SQL = "DELETE FROM RSSPool WHERE iCtNode=" & pkStr(request.queryString("ctNodeId"),"")
			conn.Execute SQL
		END IF	
		IF xUpForm("htx_YNrss") = "Y" and session("YNrss") = "N" Then
			session("ctNodeId") = request.queryString("ctNodeId")
			postURL = "/ws/ws_RSSList.asp"
			Server.Execute (postURL) 
		END IF	

	end if
	'======	
	
	'----940725目錄樹會員類別等級
	if checkGIPconfig("CtNodeMember") then
		conn.execute("Delete CtNodeMemberLvl where CtNodeID=" & pkStr(request.queryString("CtNodeID"),""))
	        if xUpForm("bfx_MType") <> "" then
		        SQLMember = ""
		    	if inStr(xUpForm("bfx_MType"),", ") = 0 then
		    	    SQLMember="Insert Into CtNodeMemberLvl values("&pkStr(request.queryString("CtNodeID"),"")&","&pkStr(xUpForm("bfx_MType"),"")&","&pkStr(xUpForm("sfx_MGrade"&xUpForm("bfx_MType")),"")&")"
		    	else
		    	    xMemberArray=split(xUpForm("bfx_MType"),", ")
			    for i=0 to ubound(xMemberArray)
			    	SQLMember=SQLMember&"Insert Into CtNodeMemberLvl values("&pkStr(request.queryString("CtNodeID"),"")&","&pkStr(xMemberArray(i),"")&","&pkStr(xUpForm("sfx_MGrade"&xMemberArray(i)),"")&");"
			    next
		    	end if
		        if SQLMember <> "" then conn.execute(SQLMember)
	        end if
	end if
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) 
	mpKey = ""
	if mpKey<>"" then  mpKey = mid(mpKey,2)
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
    <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hyweb Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>編修表單</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
'	    document.location.href="<%=doneURI%>?<%=mpKey%>"
		window.parent.Catalogue.location.reload
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>