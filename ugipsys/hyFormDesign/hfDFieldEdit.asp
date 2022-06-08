<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="ʺA]p"
HTProgFunc="s"
HTProgCode="HF011"
HTProgPrefix="htDField" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
taskLable="s" & HTProgCap

  	set htPageDom = session("hyXFormSpec")
   	set dsRoot = htPageDom.selectSingleNode("//DataSchemaDef") 	
	set xFieldNode = htPageDom.selectSingleNode("//field[fieldName='" & request("fieldName") & "']")

if request("submitTask") = "UPDATE" then

	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("Ƨs\I")
	end if

elseif request("submitTask") = "DELETE" then
	xTableName = "formUser"
	if request("part")="formList" then	xTableName="formList"
	set myParent = dsRoot.selectSingleNode("dsTable[tableName='" & xTableName &"']/fieldList")
	myParent.removeChild xFieldNode
	set session("hyXFormSpec") = htPageDom
	
	showDoneBox("ƧR\I")

else
	EditInBothCase()
end if


sub EditInBothCase
	
	showHTMLHead()
	if errMsg <> "" then	showErrBox()
	formFunction = "edit"
	showForm()
	initForm()
	showHTMLTail()
end sub


function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

	function marshItemList(fldName)
		set xn = xFieldNode  '.selectSingleNode(fldName)
		vStr = ""
		for each x in xn.selectNodes("item")
			vStr = vStr & nullText(x.selectSingleNode("mCode")) & vbCRLF
		next
		marshItemList = replace(vStr,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end function

function qqRS(fldName)
	if request("submitTask")="" then
		xValue = nullText(xFieldNode.selectSingleNode(fldName))
	else
		xValue = ""
		if request("xuf_"&fldName) <> "" then
			xValue = request("xuf_"&fldName)
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
	reg.xuf_fieldName.value= "<%=qqRS("fieldName")%>"
	reg.org_fieldName.value= "<%=qqRS("fieldName")%>"
	reg.xuf_fieldLabel.value= "<%=qqRS("fieldLabel")%>"
	reg.xuf_dataType.value= "<%=qqRS("dataType")%>"
	reg.xuf_inputType.value= "<%=qqRS("inputType")%>"
	reg.xuf_inputLen.value= "<%=qqRS("inputLen")%>"	
	initRadio "xuf_canNull", "<%=qqRS("canNull")%>"
'	reg.xuf_canNull.value= "<%=qqRS("canNull")%>"
	reg.xuf_itemList.value= "<%=marshItemList("itemList")%>"
<% if request("part")="formList" then %>
	reg.xuf_fieldDesc.value= "<%=qqRS("fieldDesc")%>"	
<% else %>
	reg.xuf_isPrimaryKey.value = "<%=qqRS("isPrimaryKey")%>"
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

    sub initOtherCheckbox(xname,ckValue,otherName)
    	valueArray = split(ckValue,",")
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
 </script>	
<%End sub '---- initForm() ----%>


<%Sub showForm()  	'===================== Client Side Validation Put HERE =========== %>
    <script language="vbscript">
      sub formModSubmit()
    
'----- ΤSubmiteˬdXbo ApU not valid  exit sub }------------
'msg1 = "аȥguȤsvAoťաI"
'msg2 = "аȥguȤW١vAoťաI"
'
'  If reg.tfx_ClientName.value = Empty Then
'     MsgBox msg2, 64, "Sorry!"
'     settab 0
'     reg.tfx_ClientName.focus
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
  reg.submitTask.value = "UPDATE"
  reg.Submit
   end sub

   sub formDelSubmit()
         chky=msgbox("`NI"& vbcrlf & vbcrlf &"@ATwRƶܡH"& vbcrlf , 48+1, "Ъ`NII")
        if chky=vbok then
	       reg.submitTask.value = "DELETE"
	      reg.Submit
       end If
  end sub

</script>

<!--#include file="hfDFieldForm.inc"-->

<%End sub '--- showForm() ------%>

<%Sub showHTMLHead() %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;"> 
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
    <title>sת</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>i<%=HTProgFunc%>j</td>
		<td width="50%" class="FormLink" align="right">
	    </td>
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
<%End sub '--- showHTMLHead() ------%>


<%Sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
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

'---- ݸˬd{Xbo	ApUҡAɳ] errMsg="xxx"  exit sub ------
'	if Request("tfx_TaxID") <> "" Then
'	  SQL = "Select * From Client Where TaxID = N'"& Request("tfx_TaxID") &"'"
'	  set RSreg = conn.Execute(SQL)
'	  if not RSreg.EOF then
'		if trim(RSreg("ClientId")) <> request("pfx_ClientID") Then
'			errMsg = "uΤ@sv!!ЭsJȤΤ@s!"
'			exit sub
'		end if
'	  end if
'	end if

End sub '---- checkDBValid() ----

Sub doUpdateDB()

	xFieldNode.selectSingleNode("fieldName").text = request("xuf_fieldName")
	xFieldNode.selectSingleNode("fieldLabel").text = request("xuf_fieldLabel")
	xFieldNode.selectSingleNode("dataType").text = request("xuf_dataType")
	xFieldNode.selectSingleNode("inputType").text = request("xuf_inputType")
	xFieldNode.selectSingleNode("inputLen").text = request("xuf_inputLen")
	xFieldNode.selectSingleNode("dataLen").text = request("xuf_inputLen")
	xFieldNode.selectSingleNode("canNull").text = request("xuf_canNull")
	xFieldNode.selectSingleNode("fieldDesc").text = request("xuf_fieldDesc")
	xFieldNode.selectSingleNode("isPrimaryKey").text = request("xuf_isPrimaryKey")
	
	for each xn in xFieldNode.selectNodes("item")
		xFieldNode.removeChild(xn)
	next
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
	set session("hyXFormSpec") = htPageDom
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) 
	mpKey = ""
	mpKey = mpKey & "&entityID=" & request("htx_entityID")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
%>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
     <title>sת</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	document.location.href = "pickField.asp"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
