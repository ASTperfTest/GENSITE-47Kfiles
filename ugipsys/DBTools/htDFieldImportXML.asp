<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料物件欄位管理"
HTProgFunc="欄位資料匯入"
HTProgCode="HT011"
HTProgPrefix="HtDfield" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function


Dim mpKey, dpKey
Dim RSmaster, RSlist
masterRef="HtDentity"
detailRef="HtDfield"

	sqlCom = "SELECT * FROM HtDentity WHERE entityId=" & pkStr(request.queryString("entityId"),"")
	Set RSmaster = Conn.execute(sqlcom)
	mpKey = ""
	mpKey = mpKey & "&entityId=" & RSmaster("entityId")
	if mpKey<>"" then  mpKey = mid(mpKey,2)


if request("submitTask") = "IMPORT" then	
	errMsg = ""
	checkDBValid()
	if errMsg <> "" then
		EditInBothCase()
	else
		doUpdateDB()
		showDoneBox("資料更新成功！")
	end if
else

	deleteFlag=false
	SQLDeleteCheck="Select * from HtDfield WHERE entityId=" & pkStr(request.queryString("entityId"),"")
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
<table border="0" width="580" cellspacing="1" cellpadding="0" align="center">
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
<CENTER><TABLE border="0" class="bluetable" cellspacing="1" cellpadding="2" width="90%"><TR><TD class="lightbluetable" align="right"><font color="red">*</font>
資料源代碼：</TD>
<TD class="whitetablebg"><input name="htx_dbID" size="10" readonly="true" class="rdonly">
<input name="htx_entityId" type="hidden" size="4">
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
	document.all("htx_dbID").value= "<%=qqRS("dbID")%>"
	document.all("htx_entityId").value= "<%=qqRS("entityId")%>"
	document.all("htx_tableName").value= "<%=qqRS("tableName")%>"
	document.all("htx_entityDesc").value= "<%=qqRS("entityDesc")%>"
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
	sql = "UPDATE HtDentity SET "
		sql = sql & "dbID=" & pkStr(request("htx_dbID"),",")
		sql = sql & "tableName=" & pkStr(request("htx_tableName"),",")
		sql = sql & "entityDesc=" & pkStr(request("htx_entityDesc"),",")
		sql = sql & "entityURI=" & pkStr(request("htx_entityURI"),",")
	sql = left(sql,len(sql)-1) & " WHERE entityId=" & pkStr(request.queryString("entityId"),"")

'	conn.Execute(SQL)  

'on error resume next

	xSeq = 0
	xTable = request("htx_tableName")

			set fxml = server.createObject("microsoft.XMLDOM")
			fxml.async = false
			fxml.setProperty "ServerHTTPRequest", true
			LoadXML = server.mappath("/HTSDSystem/xmlSpec/" & xTable & ".xml")
			
			xv = fxml.load(LoadXML)
  			if fxml.parseError.reason <> "" then
    			Response.Write(xTable & ".xml parseError on line " &  fxml.parseError.line)
    			Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    			Response.End()
  			end if

	for each xfNode in fxml.selectNodes("//dsTable[tableName='" & xTable & "']/fieldList/field")
		xSeq = xSeq + 10
		xDataType = ""
		xoDataType = nullText(xfNode.selectSingleNode("dataType"))
		rSql = "SELECT mCode FROM codeMain WHERE codeMetaID='htDdataType' AND mValue = " & pkStr(xoDataType,"")
'		response.write rSQl & "<HR>"
		set RSr = conn.execute(rSql)
		if not RSr.eof then		xDataType = RSr("mCode")
		xInputType = nullText(xfNode.selectSingleNode("inputType"))
		xidentity = nullText(xfNode.selectSingleNode("identity"))
		xisPrimaryKey = nullText(xfNode.selectSingleNode("isPrimaryKey"))
		xSql = "INSERT INTO HtDfield(entityId, xfieldName, xfieldSeq, xfieldLabel, xDataType, xdataLen, " _
			& "xinputLen, xcanNull, xisPrimaryKey, xidentity, xInputType, " _
			& "xdefaultvalue, xrefLookup, xrows, xcols"
		xValue = ") VALUES(" _
			& dfn(request.queryString("entityId")) _
			& dfs(nullText(xfNode.selectSingleNode("fieldName"))) _
			& dfn(xSeq) _
			& dfs(nullText(xfNode.selectSingleNode("fieldLabel"))) _
			& dfs(xDataType) _
			& dfs(nullText(xfNode.selectSingleNode("dataLen"))) _
			& dfs(nullText(xfNode.selectSingleNode("inputLen"))) _
			& dfs(nullText(xfNode.selectSingleNode("canNull"))) _
			& dfs(xisPrimaryKey) _
			& dfs(xidentity) _
			& dfs(xInputType) _
			& dfs(nullText(xfNode.selectSingleNode("defaultvalue"))) _
			& dfs(nullText(xfNode.selectSingleNode("refLookup"))) _
			& dfs(nullText(xfNode.selectSingleNode("rows"))) _
			& dfs(nullText(xfNode.selectSingleNode("cols"))) 
		xSql = xSql & left(xValue,len(xValue)-1) & ")"
		response.write xSql & "<HR>"
		conn.execute xSql
	next
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
	    document.location.href= "HtDfieldList.asp?<%=mpKey%>"
     </script>
     </body>
     </html>
<%End sub '---- showDoneBox() ---- %>
<%
	xTable = RSMaster("tableName")

			set fxml = server.createObject("microsoft.XMLDOM")
			fxml.async = false
			fxml.setProperty "ServerHTTPRequest", true
			LoadXML = server.mappath("/HTSDSystem/xmlSpec/" & xTable & ".xml")
			'response.write LoadXML
			xv = fxml.load(LoadXML)
  			if fxml.parseError.reason <> "" then
    			Response.Write(xTable & ".xml parseError on line " &  fxml.parseError.line)
    			Response.Write("<BR>Reason: " &  fxml.parseError.reason)
    			Response.End()
  			end if

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

	for each xfNode in fxml.selectNodes("//dsTable[tableName='" & xTable & "']/fieldList/field")

%>
	<TD class=whitetablebg><font size=2>
<%=nullText(xfNode.selectSingleNode("fieldName"))%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=nullText(xfNode.selectSingleNode("fieldLabel"))%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=nullText(xfNode.selectSingleNode("dataType"))%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=nullText(xfNode.selectSingleNode("dataLen"))%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=nullText(xfNode.selectSingleNode("isPrimaryKey"))%>
</font></td>
	<TD class=whitetablebg><font size=2>
<%=nullText(xfNode.selectSingleNode("canNull"))%>
</font></td>
    </tr>
    <%
    next
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
