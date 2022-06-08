<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="版面管理"
HTProgFunc="新增版面"
HTUploadPath="/public/"
HTProgCode="GW1M90"
HTProgPrefix="xdmp" %>
<!--#include virtual = "/inc/server.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
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
	reg.htx_Editor.value= "<%=session("userID")%>"
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
	IF reg.htx_xdmpID.value = Empty Then 
		MsgBox replace(nMsg,"{0}","版面ID"), 64, "Sorry!"
		reg.htx_xdmpID.focus
		exit sub
	END IF
	IF blen(reg.htx_xdmpID.value) > 6 Then
		MsgBox replace(replace(lMsg,"{0}","版面ID"),"{1}","6"), 64, "Sorry!"
		reg.htx_xdmpID.focus
		exit sub
	END IF
	IF reg.htx_xdmpName.value = Empty Then 
		MsgBox replace(nMsg,"{0}","版面名稱"), 64, "Sorry!"
		reg.htx_xdmpName.focus
		exit sub
	END IF
	IF blen(reg.htx_xdmpName.value) > 50 Then
		MsgBox replace(replace(lMsg,"{0}","版面名稱"),"{1}","50"), 64, "Sorry!"
		reg.htx_xdmpName.focus
		exit sub
	END IF
	IF blen(reg.htx_Purpose.value) > 80 Then
		MsgBox replace(replace(lMsg,"{0}","用途"),"{1}","80"), 64, "Sorry!"
		reg.htx_Purpose.focus
		exit sub
	END IF
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="xdmpForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
<div id="FuncName">
	<h1><%=Title%></h1>
	<div id="Nav">
 <% if (HTProgRight and 2)=3 then %><a href="Querygroup.asp">查詢</a><% End IF %>
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A> 
	</div>
	<div id="ClearFloat"></div>
</div>
<div id="FormName">新增版面</div>
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
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
'	SQL = "Select * From Client Where ClientID = N'"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "「客戶編號」重複!!請重新建立客戶編號!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	sql = "INSERT INTO XdmpList("
	sqlValue = ") VALUES("
	IF xUpForm("htx_xdmpID") <> "" Then
		sql = sql & "xdmpId" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xdmpID"),",")
	END IF
	IF xUpForm("htx_xdmpName") <> "" Then
		sql = sql & "xdmpName" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_xdmpName"),",")
	END IF
	IF xUpForm("htx_Editor") <> "" Then
		sql = sql & "editor" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_Editor"),",")
	END IF
	IF xUpForm("htx_Purpose") <> "" Then
		sql = sql & "purpose" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_Purpose"),",")
	END IF
	IF xUpForm("htx_deptID") <> "" Then
		sql = sql & "deptId" & ","
		sqlValue = sqlValue & pkStr(xUpForm("htx_deptID"),",")
	END IF


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
	set RSx = conn.Execute(SQL)
	
	xdmpID = xUpForm("htx_xdmpID")
   	Set fso = server.CreateObject("Scripting.FileSystemObject")
	filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/xdmp" & xdmpID & ".xml") 	
    if not fso.FileExists(filePath) then
		set oxml = server.createObject("microsoft.XMLDOM")
		oxml.async = false
		LoadXML = server.MapPath("/GipDSD/xmlSpec/0xdmp.xml")
		'response.write LoadXML
		xv = oxml.load(LoadXML)
		if oxml.parseError.reason <> "" then
			Response.Write("XML parseError on line " &  oxml.parseError.line)
			Response.Write("<BR>Reason: " &  oxml.parseError.reason)
			Response.End()
		end if 
		oxml.save(server.MapPath("/site/" & session("mySiteID") & "/GipDSD/xdmp" & xdmpID & ".xml"))
	end if
	
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
	document.location.href = "<%=doneURI%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
