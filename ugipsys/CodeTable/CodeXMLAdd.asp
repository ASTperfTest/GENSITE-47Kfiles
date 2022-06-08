<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="NX@"
HTProgFunc="s"
HTUploadPath="/"
HTProgCode="Pn90M02"
HTProgPrefix="CodeXML" %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
Dim Sql, SqlValue
taskLable="sW" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

function xUpForm(xvar)
	xUpForm = request.form(xvar)
end function

  	set htPageDom = session("codeXMLSpec")
  	set refModel = htPageDom.selectSingleNode("//dsTable")

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

sub clientInitForm	'---- sWɪw]ȩbo
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
    
'----- ΤSubmiteˬdXbo ApU not valid  exit sub }------------
'msg1 = "аȥguȤsvAoťաI"
'msg2 = "аȥguȤW١vAoťաI"
'
'  If reg.pfx_ClientID.value = Empty Then
'     MsgBox msg1, 64, "Sorry!"
'     settab 0
'     reg.pfx_ClientID.focus
'     Exit Sub
'  End if
	nMsg = "аȥgu{0}vAoťաI"
	lMsg = "u{0}v׳̦h{1}I"
	dMsg = "u{0}v yyyy/mm/dd I"
	iMsg = "u{0}vƭȡI"
	pMsg = "u{0}v " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 䤤@ءI"
<%
	for each param in refModel.selectNodes("fieldList/field[formList='Y']") 
		processValid param
	next
%>
  
  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="CodeXMLForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;
	    <font size=2>isW--NXID:<%=session("CodeID")%>/NXW:<%=nullText(htPageDom.selectSingleNode("//tableDesc"))%>j</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <%if (HTProgRight and 2)=2 then%>
	       <a href="<%=HTprogPrefix%>Query.asp">d</a>	    
	       <%end if%>
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
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

'---- ݸˬd{Xbo	ApUҡAɳ] errMsg="xxx"  exit sub ------
'	SQL = "Select * From Client Where ClientID = N'"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "uȤsv!!Эsإ߫Ȥs!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
	sql = "INSERT INTO  " & nullText(refModel.selectSingleNode("tableName")) & "("
	sqlValue = ") VALUES("
	for each param in refModel.selectNodes("fieldList/field[formList='Y']") 
		processInsert param
	next

	sql = left(sql,len(sql)-1) & left(sqlValue,len(sqlValue)-1) & ")"

'	response.write sql
'	response.end
	conn.Execute(SQL)  
	'----MMOSiteBz
	if nullText(refModel.selectSingleNode("tableName"))="MMOSite" then
		'----sWؿ
		Set fso = server.CreateObject("Scripting.FileSystemObject")
		fldrPath = session("Public") & request("htx_MMOSiteID")
		if not fso.FolderExists(server.MapPath(fldrPath)) then
			Set f = fso.CreateFolder(server.MapPath(fldrPath))
		end if
		'----sڥؿƦMMOFolder
		SQLI = "Insert Into MMOFolder values('/'," & _
			pkstr(request("htx_MMOSiteName"),"") & ",null," & _
			pkstr(request("htx_MMOSiteID"),"") & ",null,'zzz')"
		conn.execute(SQLI)		
		'----YFTP]w,PBsWؿ
		if request("htx_UpLoadSiteFTPIP")<>"" and request("htx_UpLoadSiteFTPPort")<>"" and request("htx_UpLoadSiteFTPID")<>"" and request("htx_UpLoadSiteFTPPWD")<>"" then
			fileAction="CreateDir"
			FTPfilePath="public"
			ftpDo request("htx_UpLoadSiteFTPIP"),request("htx_UpLoadSiteFTPPort"),request("htx_UpLoadSiteFTPID"),request("htx_UpLoadSiteFTPPWD"),fileAction,FTPfilePath,request("htx_MMOSiteID"),"","" 			
		end if
	end if
end sub '---- doUpdateDB() ----

%>

<% sub showDoneBox() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW{</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
<script language=vbs>
	alert("sWI")
	document.location.href = "<%=HTprogPrefix%>List.asp"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
