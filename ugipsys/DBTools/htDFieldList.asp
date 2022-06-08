<%@ CodePage = 65001 %>
<% 
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="資料物件欄位管理"
HTProgFunc="欄位資料"
HTProgCode="HT011"
HTProgPrefix="HtDfield" %>
<!--#include virtual = "/inc/server.inc" -->
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim mpKey, dpKey
Dim RSmaster, RSlist
masterRef="HtDentity"
detailRef="HtDfield"
if request("submitTask") = "UPDATE" then	
	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
		response.end
	end if
elseif request("submitTask") = "DELETE" then
	SQL = "DELETE FROM HtDentity WHERE entityId=" & pkStr(request.queryString("entityId"),"")
	conn.Execute SQL
	showDoneBox("資料刪除成功！")
	response.end
else

	sqlCom = "SELECT * FROM HtDentity WHERE entityId=" & pkStr(request.queryString("entityId"),"")
	Set RSmaster = Conn.execute(sqlcom)
	mpKey = ""
	mpKey = mpKey & "&entityId=" & RSmaster("entityId")
	if mpKey<>"" then  mpKey = mid(mpKey,2)
	deleteFlag=false
	SQLDeleteCheck="Select * from HtDfield WHERE entityId=" & pkStr(request.queryString("entityId"),"")
	SET RSDeleteCheck=conn.execute(SQLDeleteCheck)
	if RSDeleteCheck.EOF then deleteFlag=true
End if
%>
<body>
<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<table border="0" width="580" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>i<%=HTProgFunc%>j</td>
		<td width="50%" class="FormLink" align="right">
			<A href="htDfieldImport.asp?<%=mpKey%>">匯入db欄位</A>
			<A href="htDentityEdit.asp?<%=mpKey%>">回上一頁</A>
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
</table>
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="lightbluetable" align="right"><font color="red">*</font>
資料源代碼：</TD>
<TD class="whitetablebg"><input name="htx_dbID" size="10" readonly="true" class="rdonly">
<input name="htx_entityID" type="hidden" size="4">
</TD>
<TD class="lightbluetable" align="right"><font color="red">*</font>
資料表名稱：</TD>
<TD class="whitetablebg"><input name="htx_tableName" size="20">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">說明：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_entityDesc" size="50">
</TD>
</TR>
<TR><TD class="lightbluetable" align="right">資料來源URI：</TD>
<TD class="whitetablebg" colspan="3"><input name="htx_entityURI" size="50">
</TD>
</TR>
</TABLE>
</CENTER>
<BR>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 and deleteFlag then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
 </td></tr>
</table>
<script language=vbs>   
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
<%
function qqRS(fldName)
	xValue = RSmaster(fldname)
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function %>

<script language=vbs>
sub resetForm 
       reg.reset()
       clientInitForm
end sub
      
sub clientInitForm      
	document.all("htx_dbID").value= "<%=qqRS("dbID")%>"
	document.all("htx_entityId").value= "<%=qqRS("entityId")%>"
	document.all("htx_tableName").value= "<%=qqRS("tableName")%>"
	document.all("htx_entityDesc").value= "<%=qqRS("entityDesc")%>"
	document.all("htx_entityURI").value= "<%=qqRS("entityURI")%>"
end sub

sub window_onLoad
	clientInitForm
end sub

sub formModSubmit()    
'----- 用戶端Submit前的檢查碼放在這裡 
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"
	IF reg.htx_dbID.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料源代碼"), 64, "Sorry!"
		reg.htx_dbID.focus
		exit sub
	END IF
	IF reg.htx_entityID.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料物件代碼"), 64, "Sorry!"
		reg.htx_entityID.focus
		exit sub
	END IF
	IF reg.htx_tableName.value = Empty Then
		MsgBox replace(nMsg,"{0}","資料表名稱"), 64, "Sorry!"
		reg.htx_tableName.focus
		exit sub
	END IF
	IF blen(reg.htx_tableName.value) > 20 Then
		MsgBox replace(replace(lMsg,"{0}","資料表名稱"),"{1}","20"), 64, "Sorry!"
		reg.htx_tableName.focus
		exit sub
	END IF
	IF blen(reg.htx_entityDesc.value) > 50 Then
		MsgBox replace(replace(lMsg,"{0}","說明"),"{1}","50"), 64, "Sorry!"
		reg.htx_entityDesc.focus
		exit sub
	END IF
	IF blen(reg.htx_entityURI.value) > 100 Then
		MsgBox replace(replace(lMsg,"{0}","資料來源URI"),"{1}","100"), 64, "Sorry!"
		reg.htx_entityURI.focus
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
	sql = "UPDATE HtDentity SET "
		sql = sql & "dbId=" & pkStr(request("htx_dbId"),",")
		sql = sql & "tableName=" & pkStr(request("htx_tableName"),",")
		sql = sql & "entityDesc=" & pkStr(request("htx_entityDesc"),",")
		sql = sql & "entityUri=" & pkStr(request("htx_entityUri"),",")
	sql = left(sql,len(sql)-1) & " WHERE entityId=" & pkStr(request.queryString("entityId"),"")

	conn.Execute(SQL)  
End sub '---- doUpdateDB() ----%>

<%Sub showDoneBox(lMsg) %>
     <html>
     <head>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
     <meta name="GENERATOR" content="Hometown Code Generator 1.0">
     <meta name="ProgId" content="<%=HTprogPrefix%>List.asp">
     <title>編修表單</title>
     </head>
     <body>
     <script language=vbs>
       	alert("<%=lMsg%>")
	    document.location.href="<%=masterRef%>List.asp"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
<%
	fSql = "SELECT htx.*, xref1.mvalue AS xrxdataType FROM (HtDfield AS htx LEFT JOIN CodeMain AS xref1 ON xref1.mcode = htx.xdataType AND xref1.codeMetaId=N'htDdataType')"
	fSql = fSql & " WHERE 1=1"
	fSql = fSql & " AND htx.entityId=" & pkStr(RSmaster("entityId"),"")
	fSql = fSql & " ORDER BY htx.xfieldSeq"
	set RSlist = conn.execute(fSql)
%>
	<INPUT TYPE=button VALUE="新增資料欄位" onClick="window.navigate('htDFieldAdd.asp?<%=mpKey%>')">
	<INPUT TYPE=button VALUE="匯入db欄位" onClick="window.navigate('htDfieldImport.asp?<%=mpKey%>')">
	<INPUT TYPE=button VALUE="由XML匯入" onClick="window.navigate('htDfieldImportXML.asp?<%=mpKey%>')">
	<INPUT TYPE=button VALUE="產生XML schema" onClick="window.open('htDfieldXML.asp?<%=mpKey%>')">
	<INPUT TYPE=button VALUE="產生SQL Script" onClick="window.open('htDfieldSQL.asp?xml=<%=RSmaster("tableName")%>')">
	       <%if (HTProgRight and 4)=4 then%>
	<INPUT TYPE=button VALUE="拷貝資料物件" onClick="window.open('htDEntityCopy.asp?<%=mpKey%>')">
	       <%end if%>
	<INPUT TYPE=button VALUE="Report" onClick="window.open('htDfieldReport.asp?xml=<%=RSmaster("tableName")%>')">
	<INPUT TYPE=button VALUE="關聯表" onClick="window.open('htDFPagrp.asp?xml=<%=RSmaster("tableName")%>&xName=<%=RSmaster("entityDesc")%>')">
	       <%if (HTProgRight and 4)=4 then%>
	<INPUT TYPE=button VALUE="資料Script" onClick="window.open('htDDataSQL.asp?tbl=<%=RSmaster("tableName")%>')">
	       <%end if%>
<CENTER>
 <TABLE width="95%" cellspacing="1" cellpadding="0" class="bg">
 <tr align="left">
	<td class=lightbluetable>Seq</td>
	<td class=lightbluetable>欄位名稱</td>
	<td class=lightbluetable>標題</td>
	<td class=lightbluetable>資料型別</td>
	<td class=lightbluetable>資料長度</td>
	<td class=lightbluetable>鍵值</td>
	<td class=lightbluetable>Refer</td>
 </tr>
<%
	while not RSlist.eof
		dpKey = ""
		dpKey = dpKey & "&entityId=" & RSlist("entityId")
		dpKey = dpKey & "&xfieldName=" & RSlist("xfieldName")
		if dpKey<>"" then  dpKey = mid(dpKey,2)
%>
	<TD class=whitetablebg><font size=2>
<%=RSlist("xfieldSeq")%>
</font></td>
	<TD class=whitetablebg><font size=2>
	<A href="HtDfieldEdit.asp?<%=dpKey%>">
<%=RSlist("xfieldName")%>
</A>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("xfieldLabel")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("xrxdataType")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("xdataLen")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%if RSlist("xisPrimaryKey")="Y" then	response.write "Y"%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("xRefLookup")%>
</font></td>
    </tr>
    <%
         RSlist.moveNext
     wend
   %>
    </TABLE>
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
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

	sub butAction(k) 
	  if gpKey = "" then 
	    alert("請先選擇個案!") 
	  else
    	select case k 
    	end select 
	  end if 
	end sub 

</script>
