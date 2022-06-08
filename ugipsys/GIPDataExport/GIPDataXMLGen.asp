<%@ codepage=65001 %>
<%
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title></title>
</head>
<body>
GIP資料匯出XML檔案工具<hr>
<form id="Form1" method="POST" name="reg" action="GIPDataXMLGenList.asp">
<TABLE cellspacing="0">
<TR><TD class="eTableLable" align="right">單元資料定義：</TD>
<TD class="eTableContent"><Select name="htx_iBaseDSD" size=1>
<option value="">請選擇</option>
			<%SQL="Select iBaseDSD,sBaseDSDName from BaseDSD where iBaseDSD IS NOT NULL AND inUse='Y' Order by iBaseDSD"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
主題單元：</TD>
<TD class="eTableContent"><Select name="htx_CtUnitID" size=1>
<option value="">請選擇</option>
</select>
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
</TD>
<TD class="eTableContent">
    <input type=submit value ="顯示清單" class="cbutton">
    <input type=reset value ="重    填" class="cbutton">    
</TD>
</TR>
</table>
</form>   
</body>
</html>
<script language=vbs>
sub htx_iBaseDSD_onChange
 	set xsrc = document.all("htx_CtUnitID")
	removeOption(xsrc)
	xaddOption xsrc, "請選擇", ""

 	set oXML = createObject("Microsoft.XMLDOM")
 	oXML.async = false
 	xURI = "ws_CtUnitID.asp?iBaseDSD=" & reg.htx_iBaseDSD.value
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
