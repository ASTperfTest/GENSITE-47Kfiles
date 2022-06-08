<% Response.Expires = 0
HTProgCap="系統參數維護"
HTProgFunc="編修系統參數"
HTUploadPath="/"
HTProgCode="GW1M91"
HTProgPrefix="sysPara" %>
<!--#Include VIRTUAL = "/inc/client.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/form.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title><%=Title%></title>
</head>
<!--#INCLUDE VIRTUAL="/inc/dbutil.inc" -->
<!--#Include VIRTUAL = "/HTSDSystem/HT2CodeGen/htUIGen.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="編輯" & HTProgCap
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"

function xUpForm(xvar)
	xUpForm = request.form(xvar)
end function

  	set htPageDom = session("hyXFormSpec")
   	set dFieldSpec = htPageDom.selectSingleNode("//field[fieldName='" & request("fname") & "']")


if xUpForm("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
	end if

else
	EditInBothCase()
end if


sub EditInBothCase

	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
'	response.write sqlcom
	showForm()
	initForm()
	showHTMLTail()
end sub


function qqRS(fldName)
	if request("submitTask")="" then
		xValue = nullText(dFieldSpec.selectSingleNode(fldName))
'		xValue = dFieldSpec.selectSingleNode(fldName).xml
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
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
<%
	for each param in dFieldSpec.selectNodes("*") 
		response.write "reg.htx_" & param.nodeName & ".value= """  & qqRS(param.nodeName) & """" & vbCRLF
	next
%>
    end sub

    sub window_onLoad
         clientInitForm
    end sub
    

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
  
<%
'	for each param in refModel.selectNodes("fieldList/field[formList='Y']") 
'		processValid param
'	next
%>
  
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub


</script>

<form id="Form1" name="reg" method="POST" action="">
	<INPUT TYPE=hidden name=submitTask value="">
  <table cellspacing="0">
<%	
	for each param in dFieldSpec.selectNodes("*")
		response.write "<TR><TD class=""Label"" align=""right"">"
		if nullText(param.selectSingleNode("@must"))="Y" then _
			response.write "<span class=""Must"">*</span>"
		response.write "&lt;" & param.nodeName & "&gt;：</TD>"
		response.write "<TD class=""eTableContent"">　"
		response.write "<INPUT name=""htx_" & param.nodeName & """ size=20>　　" _
			& nullText(param.selectSingleNode("@xdesc")) 
'			processParamField param
		response.write "</TD></TR>"
	next

%>
</TABLE>
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
   <% if (HTProgRight and 8)=8 then %><input type="button" value="編修存檔" name="Enter" class="cbutton" OnClick="formModSubmit()"><% End IF %> 
   <input type=reset class=cbutton value="清除重填" onClick="resetForm()">
</form>     

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
    <body>

<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
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
	for each param in htPageDom.selectNodes("SystemParameter/*")
		param.text = request("htx_" & param.nodeName)
	next
	
	htPageDom.save(LoadXML)

End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
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
'	    document.location.href="<%=HTprogPrefix%>List.asp?<%=Session("QueryPage_No")%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
<script language=vbs>
Dim CanTarget

sub popCalendar(dateName)        
 	CanTarget=dateName
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
	end if
end sub   
</script>
