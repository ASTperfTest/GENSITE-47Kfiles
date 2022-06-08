<%@ CodePage = 65001 %>
<%
Response.Expires = 0
HTProgCap="單元資料定義"
HTProgFunc="編修"
HTProgCode="GE1T01"
HTUploadPath="/"
HTProgPrefix="BaseDsd" %>
<!--#include virtual = "/inc/server.inc" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 2.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Edit.asp">
<title>編修表單</title>
<link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim pKey
Dim RSreg
Dim formFunction
Dim mpKey, dpKey
Dim RSmaster, RSlist
taskLable="編輯" & HTProgCap

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
'		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
		Response.End
	end if

elseif xUpForm("submitTask") = "DELETE" then
	SQL = "DELETE FROM BaseDsd WHERE ibaseDsd=" & pkStr(request.queryString("ibaseDsd"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")
	Response.End
end if

	sqlCom = "SELECT htx.*,htx.inUse minUse, dhtx.xfieldName, dhtx.xfieldSeq, dhtx.inUse, dhtx.xfieldLabel, dhtx.xfieldDesc, dhtx.xdataType, dhtx.xdataLen, dhtx.xinputType, dhtx.ibaseField FROM (BaseDsd AS htx LEFT JOIN BaseDsdfield AS dhtx ON dhtx.ibaseDsd = htx.ibaseDsd) WHERE htx.ibaseDsd=" & pkStr(request.queryString("ibaseDsd"),"")	
	Set RSmaster = Conn.execute(sqlcom)	
	xTable = "CuDTx"+cStr(RSmaster("ibaseDsd"))
	mpKey = ""
	mpKey = mpKey & "&ibaseDsd=" & RSmaster("ibaseDsd")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
	pKey = mpKey
	fSql = "SELECT dhtx.*, xref1.mvalue AS inUse, xref2.mvalue AS xdataType, xref3.mvalue AS xinputType" _
		& " FROM (((BaseDsdfield AS dhtx LEFT JOIN CodeMain AS xref1 ON xref1.mcode = dhtx.inUse AND xref1.codeMetaId='boolYN') LEFT JOIN CodeMain AS xref2 ON xref2.mcode = dhtx.xdataType AND xref2.codeMetaId='htDdataType') LEFT JOIN CodeMain AS xref3 ON xref3.mcode = dhtx.xinputType AND xref3.codeMetaId='htDinputType')" _
		& " WHERE 1=1" _
		& " AND dhtx.ibaseDsd=" & pkStr(RSmaster("ibaseDsd"),"") _
		& " Order by xfieldSeq"
	set RSlist = conn.execute(fSql)
	deleteFlag=false
	if RSlist.EOF then deleteFlag=true
%>
<body>
<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	     <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 8)=8 then%>
			<A href="BaseDsdDTD.asp?<%=pKey%>" title="DTD定義">DTD</A>
		<%end if%>
		<%if (HTProgRight and 1)=1 then%>
			<A href="BaseDsdQuery.asp" title="重設查詢條件">查詢</A>
		<%end if%>
			<A href="BaseDsdList.asp" title="回單元資料清單">回前頁</A>
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
<form method="POST" name="reg" ENCTYPE="MULTIPART/FORM-DATA" action="BaseDsdEditLIst.asp?iBaseDSD=<% =request.queryString("ibaseDsd") %>">
<INPUT TYPE=hidden name=submitTask value="">
</table>
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
資料單元名稱：</TD>
<TD class="eTableContent"><input name="htx_sbaseDsdname" size="40">
</TD>
<TD class="eTableLable" align="right">資料表名稱：</TD>
<TD class="eTableContent"><input name="htx_sbaseTableName" size="30" readonly="true" class="rdonly">
<input type="hidden" name="htx_ibaseDsd">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">說明：</TD>
<TD class="eTableContent" colspan="3"><input name="htx_tdesc" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">資料單元類別：</TD>
<TD class="eTableContent"><Select name="htx_rdsdcat" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='refDSDCat' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
<TD class="eTableLable" align="right"><font color="red">*</font>
是否生效：</TD>
<TD class="eTableContent"><Select name="htx_inUse" size=1>
<option value="">請選擇</option>
			<%SQL="Select mcode,mvalue from CodeMain where msortValue IS NOT NULL AND codeMetaId='boolYN' Order by msortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

</TD>
</TR>
</TABLE>
</CENTER>
</form>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 AND deleteFlag then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">

 </td></tr>
</table>
<%
function qqRS(fldName)
	if xUpForm("submitTask")="" then
		xValue = RSmaster(fldname)
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

<script language=vbs>
sub BaseDsdUpdate()
	window.open "BaseDsdtableUpdate.asp?<%=pKey%>"
end sub
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
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

      sub resetForm 
	       reg.reset()
	       clientInitForm
      end sub

    sub window_onLoad
         clientInitForm
    end sub

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
	IF reg.htx_sbaseDsdname.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料單元名稱"), 64, "Sorry!"
		reg.htx_sbaseDsdname.focus
		exit sub
	END IF
	IF blen(reg.htx_sbaseDsdname.value) > 40 Then
		MsgBox replace(replace(lMsg,"{0}","資料單元名稱"),"{1}","40"), 64, "Sorry!"
		reg.htx_sbaseDsdname.focus
		exit sub
	END IF
	IF (reg.htx_ibaseDsd.value <> "") AND (NOT isNumeric(reg.htx_ibaseDsd.value)) Then
		MsgBox replace(iMsg,"{0}","單元資料定義ID"), 64, "Sorry!"
		reg.htx_ibaseDsd.focus
		exit sub
	END IF
	IF blen(reg.htx_tdesc.value) > 300 Then
		MsgBox replace(replace(lMsg,"{0}","說明"),"{1}","300"), 64, "Sorry!"
		reg.htx_tdesc.focus
		exit sub
	END IF
	IF reg.htx_inUse.value = Empty Then 
		MsgBox replace(nMsg,"{0}","是否生效"), 64, "Sorry!"
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

     sub clientInitForm
	reg.htx_sbaseDsdname.value= "<%=qqRS("sbaseDsdname")%>"
	reg.htx_sbaseTableName.value= "<%=qqRS("sbaseTableName")%>"
	reg.htx_ibaseDsd.value= "<%=qqRS("ibaseDsd")%>"
	reg.htx_tdesc.value= "<%=qqRS("tdesc")%>"
	reg.htx_rdsdcat.value= "<%=qqRS("rdsdcat")%>"
	reg.htx_inUse.value= "<%=qqRS("minUse")%>"
	end sub
</script>

<SCRIPT LANGUAGE="vbs">

</SCRIPT>
	<INPUT TYPE=button VALUE="新增欄位" onClick="window.navigate('bDSDFieldAdd.asp?<%=mpKey%>&phase=add')">
	<INPUT TYPE=button VALUE="匯入db欄位" onClick="window.navigate('BaseDsdFieldImport.asp?<%=mpKey%>')">
	<INPUT TYPE=button VALUE="由XML匯入" onClick="VBS: getXMLName">
	<INPUT TYPE=button VALUE="產生SQL Script" onClick="window.open('BaseDsdfieldSQL.asp?xml=<%=xTable%>&ibaseDsd=<%=RSmaster("ibaseDsd")%>')">
			
<CENTER>
 <TABLE width="95%" cellspacing="1" cellpadding="0" class="bg">
 <tr align="left">
	<td class=eTableLable>顯示順序</td>
	<td class=eTableLable>生效</td>
	<td class=eTableLable>標題</td>
	<td class=eTableLable>說明</td>
	<td class=eTableLable>資料型別</td>
	<td class=eTableLable>資料長度</td>
	<td class=eTableLable>輸入型式</td>
	<td class=eTableLable>欄位名稱</td>
 </tr>
<%
	while not RSlist.eof
		dpKey = ""
		dpKey = dpKey & "&iBaseField=" & RSlist("iBaseField")
		if dpKey<>"" then  dpKey = mid(dpKey,2)
%>
	<TD class=eTableContent><font size=2>
<%=RSlist("xfieldSeq")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("inUse")%>
</font></td>
	<TD class=eTableContent><font size=2>
	<A href="bDSDFieldEdit.asp?<%=dpKey%>&phase=edit">
<%=RSlist("xfieldLabel")%>
</A>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xfieldDesc")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xdataType")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xdataLen")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xinputType")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xfieldName")%>
</font></td>
    </tr>
    <%
         RSlist.moveNext
     wend
   %>
    </TABLE>
    </CENTER>
       <!-- { ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</body>
</html>                                 

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

sub getXMLName()
	xName = InputBox("請輸入匯入之XML檔全名:") 
	if xName="" then	exit sub
	window.navigate "BaseDsdFieldImportXML.asp?<%=mpKey%>&xmlFile="& Escape(xName)
end sub
</script>
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
	sql = "UPDATE BaseDsd SET "
		sql = sql & "sbaseDsdname=" & pkStr(xUpForm("htx_sbaseDsdname"),",")
		sql = sql & "sbaseTableName=" & pkStr(xUpForm("htx_sbaseTableName"),",")
		sql = sql & "tdesc=" & pkStr(xUpForm("htx_tdesc"),",")
		sql = sql & "rdsdcat=" & pkStr(xUpForm("htx_rdsdcat"),",")
		sql = sql & "inUse=" & pkStr(xUpForm("htx_inUse"),",")
	sql = left(sql,len(sql)-1) & " WHERE ibaseDsd=" & pkStr(request.queryString("ibaseDsd"),"")
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
	    document.location.href="<%=doneURI%>?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
