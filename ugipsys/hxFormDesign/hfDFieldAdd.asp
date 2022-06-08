<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="ʺA]p"
HTProgFunc="sW"
HTUploadPath="/"
HTProgCode="HF011"
HTProgPrefix="hfDField" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>sW</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
if request("submitTask") = "ADD" then

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
	reg.xuf_inputType.value= "<%=request("fkind")%>"
	reg.xuf_dataType.value= "varchar"
	initRadio "xuf_canNull", "Y"
<% if request("part")<>"formList" then %>
	reg.xuf_isPrimaryKey.value = "Y"
<% end if %>
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
	IF reg.xuf_fieldName.value = Empty Then
		MsgBox replace(nMsg,"{0}","W"), 64, "Sorry!"
		reg.xuf_fieldName.focus
		exit sub
	END IF
	IF reg.xuf_fieldLabel.value = Empty Then
		MsgBox replace(nMsg,"{0}","D"), 64, "Sorry!"
		reg.xuf_fieldLabel.focus
		exit sub
	END IF
	IF (reg.xuf_inputLen.value <> "") AND (NOT isNumeric(reg.xuf_inputLen.value)) Then
		MsgBox replace(iMsg,"{0}",""), 64, "Sorry!"
		reg.xuf_inputLen.focus
		exit sub
	END IF	
	IF reg.xuf_dataType.value = Empty Then
		MsgBox replace(nMsg,"{0}","ƫO"), 64, "Sorry!"
		reg.xuf_dataType.focus
		exit sub
	END IF
	IF reg.xuf_inputType.value = Empty Then
		MsgBox replace(nMsg,"{0}","J覡"), 64, "Sorry!"
		reg.xuf_inputType.focus
		exit sub
	END IF

  reg.submitTask.value = "ADD"
  reg.Submit
End Sub

</script>

<!--#include file="hfDFieldForm.inc"-->
                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>i<%=HTProgFunc%>j</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <a href="Javascript:window.history.back();">^e</a>
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
  	set htPageDom = session("hyXFormSpec")
   	set dsRoot = htPageDom.selectSingleNode("//DataSchemaDef") 	
	set xFieldNode = dsRoot.selectSingleNode("//field").cloneNode(true)

	xFieldNode.selectSingleNode("fieldName").text = request("xuf_fieldName")
	xFieldNode.selectSingleNode("fieldLabel").text = request("xuf_fieldLabel")
	xFieldNode.selectSingleNode("dataType").text = request("xuf_dataType")
	xFieldNode.selectSingleNode("inputType").text = request("xuf_inputType")
	xFieldNode.selectSingleNode("inputLen").text = request("xuf_inputLen")
	xFieldNode.selectSingleNode("dataLen").text = request("xuf_inputLen")
	xFieldNode.selectSingleNode("canNull").text = request("xuf_canNull")
	xFieldNode.selectSingleNode("fieldDesc").text = request("xuf_fieldDesc")
	xFieldNode.selectSingleNode("isPrimaryKey").text = request("xuf_isPrimaryKey")
	
	if request("xuf_itemList")<>"" then
		itemAry = split(request("xuf_itemList"),vbCRLF)
		for i = 0 to UBound(itemAry)
			if itemAry(i)<>"" then
				set nItem = htPageDom.createElement("item")
				set nCode = htPageDom.createElement("mCode")
					nCode.text = itemAry(i)
					nItem.appendChild nCode
				set nValue = htPageDom.createElement("mValue")
					nValue.text = itemAry(i)
					nItem.appendChild nValue
				xFieldNode.appendChild nItem
			end if
		next
		
	end if
	xTableName = "formUser"
	if request("part")="formList" then	xTableName="formList"
	dsRoot.selectSingleNode("dsTable[tableName='" & xTableName &"']/fieldList").appendChild xFieldNode
	set session("hyXFormSpec") = htPageDom
'	response.write Server.MapPath(".")&"\hf_" & session("hyFormID") & ".xml"
'	htPageDom.save(Server.MapPath(".")&"\hf_" & session("hyFormID") & ".xml")

end sub

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
	document.location.href = "pickField.asp"
'	document.location.href = "<%=HTprogPrefix%>List.asp?<%=request.serverVariables("QUERY_STRING")%>"
</script>
</body>
</html>
<% end sub '---- showDoneBox() ---- %>
