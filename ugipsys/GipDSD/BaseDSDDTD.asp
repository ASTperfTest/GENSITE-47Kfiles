<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="單元資料定義"
HTProgFunc="DTD定義"
HTProgCode="GE1T01"
HTProgPrefix="CtUnitDSD" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function
'Alter Table BaseDSD Add sBaseDSDXML varchar(100) null 
'----CuDTGeneric xml
set oxml = server.createObject("microsoft.XMLDOM")
oxml.async = false
oxml.setProperty "ServerHTTPRequest", true
Set fso = server.CreateObject("Scripting.FileSystemObject")
filePath = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/CuDTx" & request.querystring("ibaseDsd") & ".xml")
fileFlag = false
if fso.FileExists(filePath) then fileFlag = true
LoadXML = server.MapPath("/GipDSD/schema0.xml")
xv = oxml.load(LoadXML)
if oxml.parseError.reason <> "" then
	Response.Write("XML parseError on line " &  oxml.parseError.line)
	Response.Write("<BR>Reason: " &  oxml.parseError.reason)
	Response.End()
end if 
set allModel = oxml.selectNodes("//fieldList/field")
'----BaseDsdfield欄位
fSql = "SELECT dhtx.*, xref1.mvalue AS inUse, xref2.mvalue AS xdataType, xref3.mvalue AS xinputType" _
	& " FROM (((BaseDsdfield AS dhtx LEFT JOIN CodeMain AS xref1 ON xref1.mcode = dhtx.inUse AND xref1.codeMetaId='boolYN') LEFT JOIN CodeMain AS xref2 ON xref2.mcode = dhtx.xdataType AND xref2.codeMetaId='htDdataType') LEFT JOIN CodeMain AS xref3 ON xref3.mcode = dhtx.xinputType AND xref3.codeMetaId='htDinputType')" _
	& " WHERE 1=1" _
	& " AND dhtx.ibaseDsd=" & request.querystring("ibaseDsd") _
	& " Order by xfieldSeq"
set RSlist = conn.execute(fSql)

%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
	<td width="50%" class="FormLink">
			<A href="javascript:window.history.back();" title="回單元資料定義清單">回前頁</A>
	</td>
	<td width="50%" class="FormLink" align="right">
   		 <input type=button class=cbutton value="編修存檔" id=button1>
        </td>   
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action="BaseDSDUpdate.asp?ibaseDsd=<%=request.querystring("ibaseDsd")%>">
 <input type=hidden name="submittask" value="">
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bluetable>                   
     <tr align=left class="lightbluetable">   
	<td align=center class=eTableLable width=7%>&nbsp;</td>
	<td class=eTableLable>顯示順序</td>
	<td class=eTableLable>標題</td>
	<td class=eTableLable>說明</td>
	<td class=eTableLable>資料型別</td>
	<td class=eTableLable>輸入型式</td>
	<td class=eTableLable>欄位名稱</td>
    </tr>  
<%
for each param in allModel  
	i=i+1  
	checkStr="" 
	xmlYN="N"
	if nullText(param.selectSingleNode("isPrimaryKey"))="Y" or nullText(param.selectSingleNode("inputType"))="hidden" or nullText(param.selectSingleNode("inputType"))="hiddenSave" or nullText(param.selectSingleNode("inputType"))="SQLDefault" or nullText(param.selectSingleNode("canNull"))="N" then 
		checkStr=" checked disabled "
		xmlYN="Y"
	else
		checkStr=" checked "
		xmlYN="N"	
	end if
%>
	<TD align=center class=eTableContent><INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y" <%=checkStr%>>
	<INPUT TYPE=hidden name="fieldSeq<%=i%>" value="<%=nullText(param.selectSingleNode("fieldSeq"))%>"></td>
	<INPUT TYPE=hidden name="xmlYN<%=i%>" value="<%=xmlYN%>"></td>
	<TD class=eTableContent><font size=2>
<%=nullText(param.selectSingleNode("fieldSeq"))%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=nullText(param.selectSingleNode("fieldLabel"))%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=nullText(param.selectSingleNode("fieldDesc"))%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=nullText(param.selectSingleNode("dataType"))%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=nullText(param.selectSingleNode("inputType"))%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=nullText(param.selectSingleNode("fieldName"))%>
</font></td>
    </tr>
    <%
next
if not RSlist.eof then
    while not RSlist.eof
    	i=i+1  
%>
	<TD align=center class=eTableContent><INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y" checked disabled>
	<TD class=eTableContent><font size=2>
<%=RSlist("xfieldSeq")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xfieldLabel")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xfieldDesc")%>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSlist("xdataType")%>
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
end if    
%>      
    </TABLE>    	                             
</table> 
</form>
</body>
</html>  
<script Language=VBScript> 
sub button1_onClick
	<%if fileFlag then%>
		chky=msgbox("注意！"& vbcrlf & vbcrlf &"　已存在DTD檔案！！"& vbcrlf &"  編修存檔後將會覆蓋原有檔案(建議請先複製原有檔案以保留原有設定值), "& vbcrlf &"  確定要覆蓋檔案嗎？"& vbcrlf , 48+1, "請注意！！")
        	if chky=vbok then
			reg.submitTask.value="UPDATE"
			reg.submit
       		end If
       	<%else%>
		reg.submitTask.value="UPDATE"
		reg.submit
	<%end if%>
end sub
</script>