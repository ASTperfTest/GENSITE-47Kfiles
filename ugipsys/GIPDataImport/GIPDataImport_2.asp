<%@ codepage=65001 %>
<% Response.Expires = 0
HTProgCap="GIP資料匯入"
HTProgFunc="新增匯入(確認)"
HTProgCode="GW1M95"
HTProgPrefix="GIPDataImport" %>
<!--#include virtual = "/inc/dbutil.inc" -->
<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
<title>查詢表單</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<%
HTProgRight=255
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	Set fso = server.CreateObject("Scripting.FileSystemObject")
	INXMLPath = session("public")+"GIPDataXML/INXML"
	XMLFile = request("htx_XMLFile")
	if not fso.FileExists(server.mappath(INXMLPath)+"\"+XMLFile) then
		response.write "XML檔案不存在!"
		response.end
	end if
	if request("htx_ImportWay") = "overwrite" then
		xImportWay = "overwrite"
		xImportWayStr = "覆寫(Overwrite,刪除原檔案匯入此主題單元中資料後新增)"
	else
		xImportWay = "append"
		xImportWayStr = "接續新增(Append,原檔案匯入此主題單元中資料不刪除)"	
	end if
	SQLC = "Select G.CtUnitName from CtUnit G where G.CtUnitID="&pkStr(request("htx_CtUnitID"),"")
	Set RSC = conn.execute(SQLC)

	showHTMLHead()
	showForm()
	initForm()
	showHTMLTail()
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm

end sub

sub window_onLoad  
     clientInitForm
end sub
</script>	
<%end sub '---- initForm() ----%>

<%Sub showForm() %>
<script language="vbscript">
sub formAddSubmit()
	chky=msgbox("注意！"& vbcrlf & vbcrlf &"　您確認匯入內容嗎？"& vbcrlf , 48+1, "請注意！！")
	if chky=vbok then
		reg.action="<%=HTprogPrefix%>_Act.asp"
		reg.Submit
	end If
end Sub
</script>

<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<input type="hidden" name="htx_iBaseDSD" value="<%=request("htx_iBaseDSD")%>">
<CENTER>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right">匯入主題單元：</TD>
<TD class="eTableContent">
<input type="hidden" name="htx_CtUnitID" value="<%=request("htx_CtUnitID")%>"><%=RSC("CtUnitName")%>
</TD></TR>
<TR><TD class="eTableLable" align="right">原檔案匯入資料記錄：</TD>
<TD class="eTableContent">
	 <TABLE width="95%" cellspacing="1" cellpadding="0" border="1" class="bg">
	 <tr align="center">
		<td class=eTableLable>檔案名稱</td>
		<td class=eTableLable>匯入日期</td>
		<td class=eTableLable>匯入人員</td>
		<td class=eTableLable>成功筆數</td>	
		<td class=eTableLable>失敗筆數</td>	
	 </tr>

<%
	SQLG = "Select G.*,C.CtUnitName from GIPDataImport G Left Join CtUnit C ON G.XMLCtUnitID=C.CtUnitID " & _
			" where G.XMLCtUnitID="&pkStr(request("htx_CtUnitID"),"")&" AND XMLFile="&pkStr(XMLFile,"")&" Order by XMLDate Desc"
	Set RSG = conn.execute(SQLG)
	if not RSG.eof then
		while not RSG.eof
%>
	<TR>
		<TD class=eTableContent><font size=2>
		<%=RSG("XMLFile")%>
		</font></td>
		<TD class=eTableContent><font size=2>
		<%=RSG("XMLDate")%>
		</font></td>
		<TD class=eTableContent><font size=2>
		<%=RSG("XMLEditor")%>
		</font></td>
		<TD class=eTableContent align="right"><font size=2>
		<%=RSG("XMLSuccess")%>
		</font></td>
		<TD class=eTableContent align="right"><font size=2>
		<%=RSG("XMLFail")%>
		</font></td>
	</TR>
<%
	         RSG.moveNext
	    wend
	end if
%>	 
	</TABLE>
</TD></TR>
<TR><TD class="eTableLable" align="right">匯入XML檔案存放路徑：</TD>
<TD class="eTableContent">
<input type="hidden" name="htx_XMLFile" value="<%=XMLFile%>"><%=apath & XMLFile%>
</TD></TR>
<TR><TD class="eTableLable" align="right">
匯入方式：</TD>
<TD class="eTableContent">
<input type="hidden" name="htx_ImportWay" value="<%=xImportWay%>"><font size=2 color="red"><%=xImportWayStr%></font>
</TD></TR>
<TR><TD class="eTableLable" align="right">
匯入日期：</TD>
<TD class="eTableContent">
<%=date()%>
</TD></TR>
<TR><TD class="eTableLable" align="right">
匯入人員：</TD>
<TD class="eTableContent">
<%=session("UserName")%>
</TD></TR>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="匯  入" class="cbutton" onClick="formAddSubmit()">
              	<input type=button value ="回前頁" class="cbutton" onClick="Javascript:window.history.back();">
          <%end if%>    
 </td></tr>
</table>

<%end sub '--- showForm() ------%>

<%sub showHTMLHead() %>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
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
	    <td align=center colspan=2 width=80% height=230 valign=top>    
<%end sub '--- showHTMLHead() ------%>

<%sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<%end sub '--- showHTMLTail() ------%>
<script language=vbs>
sub htx_iBaseDSD_onChange
 	set xsrc = document.all("htx_CtUnitID")
	removeOption(xsrc)
	xaddOption xsrc, "請選擇", ""

 	set oXML = createObject("Microsoft.XMLDOM")
 	oXML.async = false
 	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_CtUnitID.asp?iBaseDSD=" & reg.htx_iBaseDSD.value
 	oXML.load(xURI)
 	set pckItemList = oXML.selectNodes("divList/row")
 	for each pckItem in pckItemList
  		xaddOption xsrc, pckItem.selectSingleNode("mValue").text,pckItem.selectSingleNode("mCode").text
 	next
 	xsrc.selectedIndex = 0
end sub

sub xaddOption(xlist,name,value)
 	set xOption = document.createElement("OPTION")
 	xOption.text=name
 	xOption.value=value
 	xlist.add(xOption)
end sub

sub removeOption(xlist)
 	for i=xlist.options.length-1 to 0 step -1
  		xlist.options.remove(i)
 	next
 	xlist.selectedIndex = -1
end sub

</script>	
