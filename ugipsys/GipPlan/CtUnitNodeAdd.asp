<% @CodePage=65001 %>
<% Response.Expires = 0
HTProgCap="目錄樹管理"
HTProgFunc="新增節點"
HTUploadPath="/public/"
HTProgCode="GE1T21"
HTProgPrefix="CtNodeT" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/checkGIPconfig.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>新增表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
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

xctNodeKind=request.querystring("ctNodeKind")
SQLC="Select mvalue from CodeMain where codeMetaId=N'refCtNodeKind' and mcode=N'"&xctNodeKind&"'"
Set RSC = conn.execute(SQLC)
'response.write sqlC
'response.end
	dim xNewIdentity
 apath=server.mappath(HTUploadPath) & "\"
' response.write apath & "<HR>"
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
	xdataLevel = 1
	if not RS.eof then
		xdataLevel = RS("dataLevel") + 1
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
	reg.htx_ctNodeKind.value= "<%=xctNodeKind%>"
	reg.htx_ctRootId.value= "<%=request("ItemID")%>"
	reg.htx_dataParent.value= "<%=request("CatID")%>"
	reg.htx_DataLevel.value= "<%=xDataLevel%>"
	reg.htx_editUserId.value= "<%=session("userID")%>"
	reg.htx_editDate.value= "<%=date()%>"
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
    
sub btn_KMcat_onClick
	window.open "http://coaw2k.hyweb.com.tw/coa/ekm/manage_doc/report2cat22.jsp?data_base_id=DB001&id_name=htx_KMcatID&autoid_name=htx_KMautoID&nm_name=htx_KMcat&&subNode=1*RB1&display=1*none&form_name=reg&focus=&catidsInput=&anchor=",null,"height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no"
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
	IF (reg.htx_dataParent.value <> "") AND (NOT isNumeric(reg.htx_dataParent.value)) Then
		MsgBox replace(iMsg,"{0}","父代節點"), 64, "Sorry!"
		reg.htx_dataParent.focus
		exit sub
	END IF
	IF (reg.htx_ctUnitId.value <> "") AND (NOT isNumeric(reg.htx_ctUnitId.value)) Then
		MsgBox replace(iMsg,"{0}","主題單元ID"), 64, "Sorry!"
		reg.htx_ctUnitId.focus
		exit sub
	END IF
	IF (reg.htx_editDate.value <> "") AND (NOT isDate(reg.htx_editDate.value)) Then
		MsgBox replace(dMsg,"{0}","編輯日期"), 64, "Sorry!"
		reg.htx_editDate.focus
		exit sub
	END IF
	IF (reg.htx_CtNodeID.value <> "") AND (NOT isNumeric(reg.htx_CtNodeID.value)) Then
		MsgBox replace(iMsg,"{0}","目錄樹節點ID"), 64, "Sorry!"
		reg.htx_CtNodeID.focus
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
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="CtUnitNodeForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=90% height=230 valign=top>   
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
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
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "「客戶編號」重複!!請重新建立客戶編號!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	sql = "INSERT INTO CatTreeNode("
	sqlValue = ") VALUES("
	IF xUpForm("htx_ctNodeKind") <> "" Then
		sql = sql & "ctNodeKind" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ctNodeKind"),",")
	END IF
	IF xUpForm("htx_ctRootId") <> "" Then
		sql = sql & "ctRootId" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ctRootId"),",")
	END IF
	IF xUpForm("htx_dataLevel") <> "" Then
		sql = sql & "dataLevel" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_dataLevel"),",")
	END IF
	IF xUpForm("htx_dataParent") <> "" Then
		sql = sql & "dataParent" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_dataParent"),",")
	END IF
	IF xUpForm("htx_ctUnitId") <> "" Then
		sql = sql & "ctUnitId" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ctUnitId"),",")
	END IF
	IF xUpForm("htx_editDate") <> "" Then
		sql = sql & "editDate" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_editDate"),",")
	END IF
	IF xUpForm("htx_editUserId") <> "" Then
		sql = sql & "editUserId" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_editUserId"),",")
	END IF
	IF xUpForm("htx_ctNodeId") <> "" Then
		sql = sql & "ctNodeId" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_ctNodeId"),",")
	END IF
	IF xUpForm("htx_catName") <> "" Then
		sql = sql & "catName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_catName"),",")
	END IF
	IF xUpForm("htx_CatShowOrder") <> "" Then
		sql = sql & "catShowOrder" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CatShowOrder"),",")
	END IF
	IF xUpForm("htx_inUse") <> "" Then
		sql = sql & "inUse" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_inUse"),",")
	END IF
'======	2006.3.30 by Gary
	IF xUpForm("htx_YNrss") <> "" Then
		sql = sql & "YNrss" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_YNrss"),",")
	END IF	
	IF xUpForm("htx_YNquery") <> "" Then
		sql = sql & "YNquery" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_YNquery"),",")
	END IF		
'======	
	IF xUpForm("htx_CtNodeNPKind") <> "" Then
		sql = sql & "CtNodeNPKind" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_CtNodeNPKind"),",")
	END IF
	
	if goFlag then
	    IF xUpForm("htx_RSSURLID") <> "" Then
		sql = sql & "RSSURLID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_RSSURLID"),",")
	    END IF
	    IF xUpForm("htx_RSSNodeType") <> "" Then
		sql = sql & "RSSNodeType" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_RSSNodeType"),",")
	    END IF
	end if	
	
	if checkGIPconfig("KMCat") then
	
		IF xUpForm("htx_KMautoID") <> "" Then
		sql = sql & "KMautoID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_KMautoID"),",")
	    END IF
	    IF xUpForm("htx_KMCatID") <> "" Then
		sql = sql & "KMCatID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_KMCatID"),",")
	    END IF
	    IF xUpForm("htx_KMCat") <> "" Then
		sql = sql & "KMCat" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_KMCat"),",")
	    END IF
	end if
	
	if checkGIPconfig("SubjectCat") then
	    IF xUpForm("htx_SubjectID") <> "" Then
		sql = sql & "SubjectID" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_SubjectID"),",")
	    END IF
	end if

    	if checkGIPconfig("SubjectMonth") then	
		sql = sql & "SubjectMonth" & ","
		sqlValue = sqlValue & pkStr(xUpForm("bfx_SubjectMonth"),",")
    	end if	
    	
	
	
For Each Form In xup.Form
If Form.IsFile Then 
    if left(Form.Name,6) = "htImg_" then
	  ofname = Form.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(Form.Name,7) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xup.Form(Form.Name).SaveAs apath & nfname, True
  	elseif left(Form.Name,7) = "htFile_" then
	  ofname = Form.FileName
	  fnExt = ""
	  if instr(ofname, ".")>0 then	fnext=mid(ofname, instr(ofname, "."))
	  tstr = now()
	  nfname = right(year(tstr),1) & month(tstr) & day(tstr) & hour(tstr) & minute(tstr) & second(tstr) & cint(rnd(second(tstr))*100) & fnext
	  sql = sql & mid(Form.Name,8) & ","
	  sqlValue = sqlValue & pkStr(nfname,",")
  	  xup.Form(Form.Name).SaveAs apath & nfname, True
  	  xsql = "INSERT INTO imageFile(newFileName, oldFileName) VALUES(" _
  	  	& pkStr(nfname,",") & pkStr(ofname,")")
  	  conn.execute xsql
	end if
end if		
  Next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"
	conn.Execute(SQL)  
	'----940725目錄樹會員類別等級
	if checkGIPconfig("CtNodeMember") then
	    if xUpForm("bfx_MType") <> "" then
	        SQLMember = ""
	    	if inStr(xUpForm("bfx_MType"),", ") = 0 then
	    	    SQLMember="Insert Into CtNodeMemberLvl values("&pkStr(xNewIdentity,"")&","&pkStr(xUpForm("bfx_MType"),"")&","&pkStr(xUpForm("sfx_MGrade"&xUpForm("bfx_MType")),"")&")"
	    	else
	    	    xMemberArray=split(xUpForm("bfx_MType"),", ")
		    for i=0 to ubound(xMemberArray)
		    	SQLMember=SQLMember&"Insert Into CtNodeMemberLvl values("&pkStr(xNewIdentity,"")&","&pkStr(xMemberArray(i),"")&","&pkStr(xUpForm("sfx_MGrade"&xMemberArray(i)),"")&");"
		    next		    	    
	    	end if
	        if SQLMember <> "" then conn.execute(SQLMember)
	    end if
	end if	 
end sub '---- doUpdateDB() ----

sub showDoneBox()
	doneURI= ""
	if doneURI = "" then	doneURI = HTprogPrefix & "List.asp"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
