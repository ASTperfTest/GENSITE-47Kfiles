<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料定義"
HTProgFunc="欄位資料匯入(由db欄位匯入)"
HTProgCode="GE1T01"
HTProgPrefix="BaseDsd" 
' ============= Modified by Chris, 2006/08/24, 修正 UniCode版 ========================'
'		Document: 950822_智庫GIP擴充.doc (修正UniCode版對nvarchar等不work)
'  modified list:
'	加	if lcase(left(xoDataType,1))="n"	then	xoDataType = mid(xoDataType,2)
'	Sub showDoneBox	    document.location.href= "BaseDsdEditList.asp? ..." 加上 phase=edit&"
%>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim mpKey, dpKey
Dim RSmaster, RSlist
Dim deleteFlag

	sqlCom = "SELECT * FROM BaseDsd WHERE ibaseDsd=" & pkStr(request.queryString("ibaseDsd"),"")
	Set RSmaster = Conn.execute(sqlcom)
	mpKey = ""
	mpKey = mpKey & "&ibaseDsd=" & RSmaster("ibaseDsd")
	if mpKey<>"" then  mpKey = mid(mpKey,2)


if request("submitTask") = "IMPORT" then	
	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("已匯入資料表欄位資料; 回單元資料定義編修畫面後, 請將欄位標題改為中文！")
	end if
else

	deleteFlag=false
	SQLDeleteCheck="Select * from BaseDsdfield WHERE ibaseDsd=" & pkStr(request.queryString("ibaseDsd"),"")
	SET RSDeleteCheck=conn.execute(SQLDeleteCheck)
	if RSDeleteCheck.EOF then deleteFlag=true
End if
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title></title>
</head>
<body>
<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<table border="0" width="95%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
			<A href="javascript: history.back();">回上一頁</A>
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
<CENTER>
<TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right"><font color="red">*</font>
資料單元名稱：</TD>
<TD class="eTableContent"><input name="htx_sBaseDsdName" size="40">
</TD>
<TD class="eTableLable" align="right">資料表名稱：</TD>
<TD class="eTableContent"><input name="htx_sBaseTableName" size="30" readonly="true" class="rdonly">
<input type="hidden" name="htx_ibaseDsd">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">說明：</TD>
<TD class="eTableContent" colspan="3"><input name="htx_tDesc" size="50">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">資料單元類別：</TD>
<TD class="eTableContent"><Select name="htx_rDSDCat" size=1>
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
<BR>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><th width="100%">     
       <% if (HTProgRight AND 4) <> 0 and deleteFlag then %>
             <input type=button value ="匯入欄位" class="cbutton" onClick="formModSubmit()">   
       <%end If %>           
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
	reg.htx_sBaseDsdName.value= "<%=qqRS("sBaseDsdName")%>"
	reg.htx_sBaseTableName.value= "<%=qqRS("sBaseTableName")%>"
	reg.htx_ibaseDsd.value= "<%=qqRS("ibaseDsd")%>"
	reg.htx_tDesc.value= "<%=qqRS("tDesc")%>"
	reg.htx_rDSDCat.value= "<%=qqRS("rDSDCat")%>"
	reg.htx_inUse.value= "<%=qqRS("inUse")%>"
end sub

sub window_onLoad
	clientInitForm
end sub

sub formModSubmit()      
  reg.submitTask.value = "IMPORT"
  reg.Submit
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
	fSql = "sp_pkeys " & pkStr(RSmaster("sBaseTableName"),"")
	set RSlist = conn.execute(fSql)
	rsPkeys = ""
	while not RSlist.eof
		rsPkeys = rsPkeys & pkStr(RSlist("COLUMN_NAME"),",")
		RSlist.moveNext
	wend

	fSql = "sp_columns " & pkStr(RSmaster("sBaseTableName"),"")
	set RSlist = conn.execute(fSql)
	xSeq = 500
	while not RSlist.eof
		xSeq = xSeq + 10
		xDataType = ""
		xInputType = "textbox"
		xidentity = "N"
		xoDataType = RSlist("TYPE_NAME")
		if lcase(left(xoDataType,1))="n"	then	xoDataType = mid(xoDataType,2)
		xpos = inStr(xoDataType, " identity")
		if xpos > 0 then
			xoDataType = left(xoDataType,xpos-1)
			xidentity = "Y"
			xInputType = "hidden"
		end if
		rSql = "SELECT mcode FROM CodeMain WHERE codeMetaId='htDdataType' AND mvalue = " & pkStr(xoDataType,"")
		response.write rSQl & "<HR>"
		set RSr = conn.execute(rSql)
		if not RSr.eof then		xDataType = RSr("mcode")
		xisPrimaryKey = "N"
		if instr(rsPkeys, pkStr(RSlist("COLUMN_NAME"),",")) > 0 then	xisPrimaryKey = "Y"
		xdefaultvalue = RSlist("COLUMN_DEF")
		xSql = "INSERT INTO BaseDsdfield(ibaseDsd, xfieldName, xfieldSeq, xfieldLabel, xDataType, xdataLen, " _
			& "xcanNull, xisPrimaryKey, xidentity, xInputType"
		xValue = ") VALUES(" _
			& dfn(request.queryString("ibaseDsd")) _
			& dfs(RSlist("COLUMN_NAME")) _
			& dfn(xSeq) _
			& dfs(RSlist("COLUMN_NAME")) _
			& dfs(xDataType) _
			& dfs(RSlist("LENGTH")) _
			& dfs(mid(RSlist("IS_NULLABLE"),1,1)) _
			& dfs(xisPrimaryKey) _
			& dfs(xidentity) _
			& dfs(xInputType) 
		if RSlist("COLUMN_DEF") <> "" then
			xSql = xSql & ", xdefaultvalue"
			xValue = xValue & RSlist("COLUMN_DEF") & ","
		end if
		xSql = xSql & left(xValue,len(xValue)-1) & ")"
		response.write xSql & "<HR>"
		conn.execute xSql
      RSlist.moveNext
    wend

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
	    document.location.href= "BaseDsdEditList.asp?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
<%
	fSql = "sp_pkeys " & pkStr(RSmaster("sBaseTableName"),"")
	set RSlist = conn.execute(fSql)
	rsPkeys = ""
	while not RSlist.eof
		rsPkeys = rsPkeys & pkStr(RSlist("COLUMN_NAME"),",")
		RSlist.moveNext
	wend

	fSql = "sp_columns " & pkStr(RSmaster("sBaseTableName"),"")
	set RSlist = conn.execute(fSql)
%>
<CENTER>
 <TABLE width="95%" cellspacing="1" cellpadding="0" class="bg">
 <tr align="left">
	<td class=lightbluetable>欄位名稱</td>
	<td class=lightbluetable>標題</td>
	<td class=lightbluetable>資料型別</td>
	<td class=lightbluetable>資料長度</td>
	<td class=lightbluetable>鍵值</td>
	<td class=lightbluetable>Nullable</td>
 </tr>
<%
	while not RSlist.eof
		dpKey = ""
'		dpKey = dpKey & "&entityID=" & RSlist("entityID")
		if dpKey<>"" then  dpKey = mid(dpKey,2)
%>
	<TD class=whitetablebg><font size=2>
<%=RSlist("COLUMN_NAME")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("COLUMN_NAME")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("TYPE_NAME")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("LENGTH")%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%
	if instr(rsPkeys, pkStr(RSlist("COLUMN_NAME"),",")) > 0 then
		response.write "Y"
	end if
%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=RSlist("IS_NULLABLE")%>
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

