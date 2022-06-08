<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgFunc="新增從屬部門"
HTUploadPath="/public/"
HTProgCode="Pn02M02"
HTProgPrefix="dept" 
' ============= Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'
'		Document: 950822_智庫GIP擴充.doc
'  modified list:
'	存檔時檢查GIPconfig<DeptCtUnitID>，有值時另插入一筆 CuDtGeneric
' ============= Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================
'		Document: 950822_智庫GIP擴充.doc
'  modified list:
'	Function ChrToDec(inChr)
'	Function DecToChr(inDec)
'	存檔時產生新代碼
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
' ===begin========== Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================
Function ChrToDec(inChr)
dim xv
 if len(inChr) = 1 then
   temp=Asc(inChr)
   if temp > 64 then
      ChrToDec=temp-55
   else
      ChrToDec=temp-48
   end if
 else
	xv = 0
	for xi=1 to len(inChr)
		xv = xv*36 + ChrToDec(mid(inChr,xi,1))
	next
	ChrToDec = xv
 end if
End Function

Function DecToChr(inDec)
   if inDec < 10 then
      DecToChr=Cstr(inDec)
   elseif inDec < 36 then
      DecToChr=Chr(inDec+55)
   else
		xs = ""
		xv = inDec
		while xv > 0 
			xrv = xv mod 36
			xv = xv \ 36
			xs = DecToChr(xrv) & xs
		wend
		DecToChr = xs
   end if
End Function
' ===end========== Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================

	dim xNewIdentity
 apath=server.mappath(HTUploadPath) & "\"
  if request.querystring("phase")<>"add" then
Set xup = Server.CreateObject("TABS.Upload")
xup.codepage=65001
xup.Start apath

else
Set xup = Server.CreateObject("TABS.Upload")
end if

function xUpForm(xvar)
	xUpForm = xup.form(xvar)
end function

if xUpForm("submitTask") = "ADD" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		showErrBox()
	else
		doUpdateDB()
		showDoneBox()
	end if

else
	sql = "SELECT dataLevel FROM CatTreeNode WHERE ctNodeId=" & pkStr(request("CatID"),"")
	set RS = conn.execute(sql)
	xDataLevel = 1
	if not RS.eof then
		xDataLevel = RS("DataLevel") + 1
	end if


	showHTMLHead()
	formFunction = "add"
	showForm()
	initForm()
	showHTMLTail()
end if

%>
<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm	'---- 新增時的表單預設值放在這裡
	reg.htx_parent.value= "<%=request("deptID")%>"
	reg.htx_inUse.value= "Y"
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

    sub initCheckbox(xname,ckValue)
    	value = ckValue & ","
    	for each xri in reg.all(xname)
    		if instr(value, xri.value&",") > 0 then 
    			xri.checked = true
    		end if
    	next
    end sub
</script>	
<% end sub '---- initForm() ----%>

<% sub showForm() 	'===================== Client Side Validation Put HERE =========== %>
<script language="vbscript">

Sub formAddSubmit()	
    
'----- 用戶端Submit前的檢查碼放在這裡 ，如下例 not valid 時 exit sub 離開------------
'msg1 = "請務必填寫「客戶編號」，不得為空白！"
'msg2 = "請務必填寫「客戶名稱」，不得為空白！"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
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
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="deptForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
	    <a href="VBScript: history.back">回上一頁</a>
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName"><%=HTProgFunc%></div>

<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
<div id="Explain">
	<h1>說明</h1>
	<ul>
		<li><span class="Must">*</span>為必要欄位</li>
	</ul>
</div>
</body>
</html>
</html>
<% end sub '--- showHTMLTail() ------%>


<% sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
  			window.history.back
	</script>
<%
end sub '---- showErrBox() ----

sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- 後端資料驗證檢查程式碼放在這裡	，如下例，有錯時設 errMsg="xxx" 並 exit sub ------
'	SQL = "Select * From Client Where ClientID = N'"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "「客戶編號」重複!!請重新建立客戶編號!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()

' ===begin========== Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================
	dblDeptIDpad = ""
	dblDeptIDlen = 1
 	if checkGIPconfig("DeptIdDouble") then
		dblDeptIDpad = "0"
		dblDeptIDlen = 2
	end if	

	parentID = xUpForm("htx_parent")
	
	'--------將這筆的ID抓出來改
  SQLcom = "select deptId, seq from Dept where parent = " & pkstr(parentid,"") & " order by deptId Desc"
  set RS1 = conn.execute(sqlCom)
  if rs1.eof then 
		myseq = "1"
    newhidid = parentID & dblDeptIDpad & "1"
  else  
		myseq = RS1("seq") + 1
		'-------找空的deptID
    ARYChkDptID = rs1.getrows(999)
    i = 1
    For xi = ubound(ARYChkDptID, 2) to 0 step - 1
	temp = ChrToDec( right( ARYChkDptID(0, xi), dblDeptIDlen) )
	temp1 = right( ARYChkDptID(0, xi), dblDeptIDlen)
      if temp <> i and isnumeric(temp1) then
	newhidid=left(ARYChkDptID(0,xi),len(ARYChkDptID(0,xi))-dblDeptIDlen) & right("00"&DecToChr(temp-1),dblDeptIDlen)
        exit for
      else
	i = i + 1
      end if   
      tempDepID = ARYChkDptID( 0, xi)
    next
		'-------------找不到空的deptID ,新增deptID
    if newhidid = "" then  
      plen = len(tempDepID)
      if plen > 1 then
        tranpar = ChrToDec(right(tempDepID,dblDeptIDlen))
      end if
      tranpar = tranpar + 1        
      if plen > 2 then
        newhidid = left(tempDepID, plen - dblDeptIDlen) & right( "00" & DecToChr(tranpar), dblDeptIDlen )
      else
        newhidid = left(tempDepID, 1) & right( "00" & DecToChr(tranpar), dblDeptIDlen )
      end if
    end if
  end if
'response.write newhidid
'response.end
	IF xUpForm("htx_seq") <> "" Then	myseq = xUpForm("htx_seq")
' ===end========== Modified by Chris, 2006/08/22, to handle "DeptIdDouble" ========================

	sql = "INSERT INTO Dept(deptId, parent,seq,"
	sqlValue = ") VALUES(" & pkStr(newHidid,",") & pkStr(parentID,",") & pkStr(myseq,",")
' ===begin========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'
	CtUnitID = getGIPconfigText("DeptCtUnitID")
	if CtUnitID<>"" then
'		response.write CtUnitID
		xsql = "INSERT INTO CuDtGeneric(ibaseDSD,iCtUnit,sTitle, iEditor,idept,dEditDate,showType) VALUES(" _
			& "8," & pkStr(CtUnitID,",") _
			& pkStr(xUpForm("htx_deptName"),",") _
			& pkStr(session("userID"),",") _
			& pkStr(newHidid,",") _
			& "getdate(),'1')"
		xsql = "set nocount on;"&xsql&"; select @@IDENTITY as NewID"
		set RSx = conn.Execute(xsql)
		xNewIdentity = RSx(0)
		sql = sql & "giCuItem" & ","
		sqlValue = sqlValue & pkStr(xNewIdentity,",")
	end if
'	response.end
' ===end========== Modified by Chris, 2006/08/24, to handle 單位組織作為CuDtSpecific ========================'

	IF xUpForm("htx_deptName") <> "" Then
		sql = sql & "deptName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_deptName"),",")
	END IF
	IF xUpForm("htx_AbbrName") <> "" Then
		sql = sql & "abbrName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_AbbrName"),",")
	END IF
	IF xUpForm("htx_eDeptName") <> "" Then
		sql = sql & "edeptName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_eDeptName"),",")
	END IF
	IF xUpForm("htx_eAbbrName") <> "" Then
		sql = sql & "eabbrName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_eAbbrName"),",")
	END IF
  	if checkGIPconfig("dpetOrgCode") then
	IF xUpForm("htx_deptCode") <> "" Then
		sql = sql & "deptCode" & "," & "codeName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_deptCode"),",")
		sqlValue = sqlValue & pkStr(xUpForm("htx_deptName") & " (" & xUpForm("htx_deptCode") & ")",",")
	END IF
	end if
	IF xUpForm("htx_OrgRank") <> "" Then
		sql = sql & "orgRank" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_OrgRank"),",")
	END IF
	IF xUpForm("htx_kind") <> "" Then
		sql = sql & "kind" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_kind"),",")
	END IF
	IF xUpForm("htx_inUse") <> "" Then
		sql = sql & "inUse" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_inUse"),",")
	END IF
	IF xUpForm("htx_tDataCat") <> "" Then
		sql = sql & "tdataCat" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_tDataCat"),",")
	END IF
	
	IF xUpForm("htx_add") <> "" Then
		sql = sql & "servAddr" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_add"),",")
	END IF
	IF xUpForm("htx_website") <> "" Then
		sql = sql & "webSite" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_website"),",")
	END IF
	IF xUpForm("htx_tel") <> "" Then
		sql = sql & "servPhone" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_tel"),",")
	END IF

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
'	response.write sql
'	response.end
	conn.Execute(SQL)  
end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增程式</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("新增完成！")
'	document.location.href = "<%=doneURI%>"
	window.parent.Catalogue.location.reload
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
