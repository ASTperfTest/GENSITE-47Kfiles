<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgFunc="編修組織"
HTUploadPath="/public/"
HTProgCode="Pn02M02"
HTProgPrefix="dept" 
' ============= Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'
'		Document: 950822_智庫GIP擴充.doc
'  modified list:
'	更新/刪除時檢查GIPconfig<DeptCtUnitID>，有值時另更新/刪除 CuDtGeneric
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction

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
' ===begin========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'
	CtUnitID = getGIPconfigText("DeptCtUnitID")
	if CtUnitID<>"" then
		giCuItem = xUpForm("htx_giCuItem")
		if isNumeric(giCuItem) then
			SQL = "DELETE CuDtGeneric WHERE iCtUnit=" & pkStr(CtUnitID,"") _
				& " AND iCuItem=" & pkStr(giCuItem,"")
			conn.execute SQL
		end if
	end if
' ===end========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'
	SQL = "DELETE FROM Dept WHERE deptId=" & pkStr(request.queryString("deptID"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")

else
	EditInBothCase()
end if


sub EditInBothCase
	sqlSelect = "SELECT d.*, (SELECT count(*) FROM Dept WHERE parent=d.deptId) AS child "
	sqlFrom = " FROM Dept AS d "
	'WHERE d.deptId=" & pkStr(request.queryString("deptID"),"")
  	if checkGIPconfig("DeptHeadSet") then
		sqlSelect = sqlSelect & ", u.userName AS DeptHeadUser"
		sqlFrom = sqlFrom & "JOIN infoUser AS u ON u.userID=d.deptHead " 
	end if
	sqlCom = sqlSelect & sqlFrom _
		& " WHERE d.deptId=" & pkStr(request.queryString("deptID"),"")			
		
	Set RSreg = Conn.execute(sqlcom)
	pKey = ""
	pKey = pKey & "&deptID=" & RSreg("deptID")
	if pKey<>"" then  pKey = mid(pKey,2)

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if xUpForm("submitTask")="" then
		'on error resume next
		xValue = RSreg(fldname)
		
	else
		xValue = ""
		if xUpForm("htx_"&fldName) <> "" then
			xValue = xUpForm("htx_"&fldName)
		end if
	end if
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
	reg.htx_deptID.value= "<%=qqRS("deptID")%>"
	reg.htx_deptName.value= "<%=qqRS("deptName")%>"
	reg.htx_AbbrName.value= "<%=qqRS("AbbrName")%>"
	reg.htx_eDeptName.value= "<%=qqRS("eDeptName")%>"
	reg.htx_eAbbrName.value= "<%=qqRS("eAbbrName")%>"
<% 	if checkGIPconfig("dpetOrgCode") then %>
	reg.htx_deptCode.value= "<%=qqRS("deptCode")%>"
<%	end if %>
<% 	if checkGIPconfig("DeptHeadSet") then %>
	reg.htx_DeptHead.value= "<%=qqRS("DeptHead")%>"
	reg.htx_DeptHeadUser.value= "<%=qqRS("DeptHeadUser")%>"
<%	end if %>
	reg.htx_add.value= "<%=qqRS("servAddr")%>"
	reg.htx_tel.value= "<%=qqRS("servPhone")%>"
	
	reg.htx_OrgRank.value= "<%=qqRS("OrgRank")%>"	
	reg.htx_seq.value= "<%=qqRS("seq")%>"
	reg.htx_kind.value= "<%=qqRS("kind")%>"
	reg.htx_inUse.value= "<%=qqRS("inUse")%>"
	reg.htx_website.value= "<%=qqRS("website")%>"
<% 	if checkGIPconfig("DeptTopCat") then %>
	initCheckbox "htx_tDataCat","<%=qqRS("tDataCat")%>"
<%	end if %>
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
  
	IF reg.htx_deptName.value = Empty Then 
		MsgBox replace(nMsg,"{0}","部門中文名稱"), 64, "Sorry!"
		reg.htx_deptName.focus
		exit sub
	END IF
	IF blen(reg.htx_deptName.value) > 70 Then
		MsgBox replace(replace(lMsg,"{0}","部門中文名稱"),"{1}","70"), 64, "Sorry!"
		reg.htx_deptName.focus
		exit sub
	END IF
	IF blen(reg.htx_AbbrName.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","單位簡稱"),"{1}","30"), 64, "Sorry!"
		reg.htx_AbbrName.focus
		exit sub
	END IF
	IF blen(reg.htx_eDeptName.value) > 60 Then
		MsgBox replace(replace(lMsg,"{0}","英文名稱"),"{1}","60"), 64, "Sorry!"
		reg.htx_eDeptName.focus
		exit sub
	END IF
	IF blen(reg.htx_eAbbrName.value) > 30 Then
		MsgBox replace(replace(lMsg,"{0}","英文簡稱"),"{1}","30"), 64, "Sorry!"
		reg.htx_eAbbrName.focus
		exit sub
	END IF
	IF reg.htx_inUse.value = Empty Then
		MsgBox replace(nMsg,"{0}","是否有效"), 64, "Sorry!"
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

<!--#include file="deptFormE.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
	    <a href="VBScript: history.back">回上一頁</a>
	    <a href="deptAdd.asp?<%=pkey%>&phase=add">新增從屬部門</a>
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
'	  SQL = "Select * From Client Where TaxID = N'"& Request("tfx_TaxID") &"'"
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
' ===begin========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'
	CtUnitID = getGIPconfigText("DeptCtUnitID")
	if CtUnitID<>"" then
		giCuItem = xUpForm("htx_giCuItem")
		if isNumeric(giCuItem) then
			SQL = "UPDATE CuDtGeneric SET" _
				& " sTitle=" & pkStr(xUpForm("htx_deptName"),",") _
				& " iEditor=" & pkStr(session("userID"),",") _
				& " dEditDate=getdate()" 
			conn.execute SQL
		end if
	end if
' ===end========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'
	sql = "UPDATE Dept SET "
		sql = sql & "deptName=" & pkStr(xUpForm("htx_deptName"),",")
		sql = sql & "abbrName=" & pkStr(xUpForm("htx_AbbrName"),",")
		sql = sql & "edeptName=" & pkStr(xUpForm("htx_eDeptName"),",")
		sql = sql & "eabbrName=" & pkStr(xUpForm("htx_eAbbrName"),",")
   	if checkGIPconfig("DeptHeadSet") then
		sql = sql & "DeptHead=" & pkStr(xUpForm("htx_deptHead"),",")
	end if
 	if checkGIPconfig("dpetOrgCode") then
		sql = sql & "deptCode=" & pkStr(xUpForm("htx_deptCode"),",")
		sql = sql & "codeName=" & pkStr(xUpForm("htx_deptName") & " (" & xUpForm("htx_deptCode") & ")",",")
	end if
		sql = sql & "seq=" & pkStr(xUpForm("htx_seq"),",")
		sql = sql & "servAddr=" & pkStr(xUpForm("htx_add"),",")
		sql = sql & "webSite=" & pkStr(xUpForm("htx_website"),",")
		sql = sql & "servPhone=" & pkStr(xUpForm("htx_tel"),",")
		sql = sql & "orgRank=" & pkStr(xUpForm("htx_OrgRank"),",")
		sql = sql & "kind=" & pkStr(xUpForm("htx_kind"),",")
		sql = sql & "inUse=" & pkStr(xUpForm("htx_inUse"),",")
		sql = sql & "tdataCat=" & pkStr(xUpForm("htx_tDataCat"),",")
	sql = left(sql,len(sql)-1) & " WHERE deptId=" & pkStr(request.queryString("deptID"),"")

	conn.Execute(SQL)  
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
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
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
