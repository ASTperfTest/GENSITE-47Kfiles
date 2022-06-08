<%@ codepage=65001 %>
<% Response.Expires = 0
HTProgCap="GIP資料匯入"
HTProgFunc="新增匯入"
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

taskLable="查詢" & HTProgCap

	showHTMLHead()
	formFunction = "query"
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
    	 if reg.htx_iBaseDSD.value = "" then
    	 	alert "請選擇資料範本定義!"
    	 	reg.htx_iBaseDSD.focus
    	 	exit sub
    	 end if
    	 if reg.htx_CtUnitID.value = "" then
    	 	alert "請選擇主題單元!"
    	 	reg.htx_CtUnitID.focus
    	 	exit sub
    	 end if
    	 if reg.htx_XMLFile.value = "" then
    	 	alert "請輸入匯入XML檔案路徑!"
    	 	reg.htx_XMLFile.focus
    	 	exit sub
    	 end if
    	 if reg.htx_ImportWay.item(0).checked = false and reg.htx_ImportWay.item(1).checked = false then
    	 	alert "請選擇匯入方式!"
    	 	exit sub    	 
    	 end if
         reg.action="<%=HTprogPrefix%>_2.asp"
         reg.Submit
    end Sub
   </script>

<form method="POST" name="reg" action="">
<INPUT TYPE=hidden name=submitTask value="">
<CENTER>
<TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR><TD class="eTableLable" align="right">資料範本定義：</TD>
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
<TR><TD class="eTableLable" align="right">匯入XML檔案名稱：</TD>
<TD class="eTableContent"><input type="text" name="htx_XMLFile" size="30">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">
匯入方式：</TD>
<TD class="eTableContent">
<input type="radio" name="htx_ImportWay" value="overwrite" checked>覆寫(Overwrite,刪除原檔案匯入此主題單元中的資料後, 再新增)<br>
<input type="radio" name="htx_ImportWay" value="append">接續新增(Append,原檔案匯入此主題單元中的資料不刪除)
</TD>
</TR>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="下一步" class="cbutton" onClick="formAddSubmit()">
              	<input type=button value ="重　填" class="cbutton" onClick="resetForm()">
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
