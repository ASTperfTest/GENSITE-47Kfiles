<%@ CodePage = 65001 %>
<html>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
<body style="margin-left:0; margin-top:3;margin-right:0">
<table cellspacing=2 border=1>
<TR>
<TH>標題</TH>
<TH>欄位名稱</TH>
<TH>資料型別</TH>
<TH>輸入方式</TH>
<TH>長度</TH>
<TH>可空</TH>
<TH>編修</TH>
</TR>
<% 
function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

  	set htPageDom = session("hyXFormSpec")
   	set dsRoot = htPageDom.selectSingleNode("//DataSchemaDef") 	

	listField dsRoot.selectSingleNode("//dsTable[tableName='formUser']")
		response.write "<TR><TD align=left colspan=7>" 
		response.write "<A href=""hfDFieldAdd.asp"">新增欄位</A>"
		response.write "</TD></TR>"	
	if nullText(dsRoot.selectSingleNode("//dsTable[tableName='formList']/tableName"))="formList" then
	  listField dsRoot.selectSingleNode("//dsTable[tableName='formList']")
		response.write "<TR><TD align=left colspan=7>" 
		response.write "<A href=""hfDFieldAdd.asp?part=formList"">新增欄位</A>"
		response.write "　<BUTTON class=cbutton onclick=""VBS:genList 0"">產生表列</BUTTON>"
		response.write "　<BUTTON class=cbutton onclick=""VBS:genList 1"">產生表列帶項次</BUTTON>"
		response.write "</TD></TR>"	
	end if
	listField dsRoot.selectSingleNode("//dsTable[tableName='formGeneric']")
	listField dsRoot.selectSingleNode("//dsTable[tableName='hfAptForm']")

sub listField(xTable)
		response.write "<TR bgcolor=""#CCCCCC""><TH align=left colspan=7>" 
		response.write nullText(xTable.selectSingleNode("tableDesc")) 
		response.write "</TH></TR>"	
		xPart = ""
		if nullText(xTable.selectSingleNode("tableName"))="formList"	then	xPart = "&part=formList"
	for each xField in xTable.selectNodes("fieldList/field")	
		xfID = nullText(xField.selectSingleNode("fieldSeq"))
		xfName = nullText(xField.selectSingleNode("fieldName"))
		xfLabel = nullText(xField.selectSingleNode("fieldLabel"))
	
%>
<TR><TD><SPAN id="df_<%=xfName%>" title="<%=xfName%>" style="cursor:hand"><%=xfLabel%>
</SPAN></TD>
<TD><%=xfName%></TD>
<TD><%=nullText(xField.selectSingleNode("dataType"))%></TD>
<TD><%=nullText(xField.selectSingleNode("inputType"))%></TD>
<TD><%=nullText(xField.selectSingleNode("inputLen"))%></TD>
<TD><%=nullText(xField.selectSingleNode("canNull"))%></TD>
<% if nullText(xTable.selectSingleNode("tableName"))<>"hfAptForm" AND nullText(xTable.selectSingleNode("tableName"))<>"formGeneric" then %>
<TD><A href="hfDFieldEdit.asp?fieldName=<%=xfName%><%=xPart%>">編修</A></TD>
<% end if %>
</TR>
<% 		
	next
end sub
%>
</table>
</body>
<script language=vbs>
	set htPageDom = CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

sub document_onClick
	LoadXML = "ws/processparamField.asp?fname="
	set srcItem = window.event.srcElement
  if srcItem.tagName = "SPAN" then
'		msgBox srcItem.id
		grid = srcItem.id
'		msgBox parent.parent.frames("fedit").document.all.tags("SPAN").length
'		set allSpan = parent.parent.frames("fedit").document.all.tags("SPAN")
'		for each xs in allSpan
'			if xs.id = grid then
'				srcItem.outerHTML = ""
'				exit sub
'			end if
'		next
	xStyle=" id="&grid&" style=""background-color:#EEE;"">"
	xv = htPageDom.load(LoadXML&mid(grid,4))
  	if htPageDom.parseError.reason <> "" then 
    		msgBox(LoadXML&mid(grid,4) & " htPageDom parseError on line " &  htPageDom.parseError.line)
    		msgBox("Reasonxx: " &  htPageDom.parseError.reason)
    		exit sub
  	end if
	writeCodeStr = nullText(htPageDom.selectSingleNode("//pFieldHTML"))
	xHTML = "<SPAN "&xstyle&writeCodeStr&"</SPAN><BR/>"
'	call parent.parent.frames("fedit").insertField (xHTML,grid)
		window.opener.doInsertField writeCodeStr,grid
		window.close


  end if
end sub

sub genList(withNum)
	rtValue = inputbox("重覆次數, tabIndex起始值")
	xTabIndex = 100
	rtArray = split(rtValue, ",")
	nList = trim(rtArray(0))
	if uBound(rtArray) = 1 then	xTabIndex = trim(rtArray(1))
	if not isNumeric(nList)	OR nList="" then	exit sub
	LoadXML = "ws/processParamList.asp?repeatTimes="&nList&"&xtIndex="&xTabIndex&"&withNum=" & withNum
	xv = htPageDom.load(LoadXML)
  	if htPageDom.parseError.reason <> "" then 
    		msgBox( " htPageDom parseError on line " &  htPageDom.parseError.line)
    		msgBox("Reasonxx: " &  htPageDom.parseError.reason)
    		exit sub
  	end if
	writeCodeStr = nullText(htPageDom.selectSingleNode("//pFieldHTML"))
'	alert writeCodeStr
		window.opener.doInsertField writeCodeStr,grid
		window.close
end sub
</script>
</html>