<%@ CodePage = 65001 %>
<html>
<body style="margin-left:0; margin-top:3;margin-right:0">
<table cellspacing=2 border=1>
<!--#include FILE = "dhtUIGen.inc" -->
<% 

	xmlSpec = "fCuDTx9"
	set htPageDom = Server.CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false

'	LoadXML = server.MapPath("/site/" & session("mySiteID") & "/GipDSD/" & xmlSpec & ".xml")
	LoadXML = server.MapPath(xmlSpec & ".xml")
	xv = htPageDom.load(LoadXML)
'	response.write xv & "<HR>"
  	if htPageDom.parseError.reason <> "" then 
    		Response.Write("htPageDom parseError on line " &  htPageDom.parseError.line)
    		Response.Write("<BR>Reasonxx: " &  htPageDom.parseError.reason)
    		Response.End()
  	end if
  	set session("hyXFormSpec") = htPageDom
   	set dsRoot = htPageDom.selectSingleNode("DataSchemaDef") 	

	for each xField in dsRoot.selectNodes("//field")	
		xfID = nullText(xField.selectSingleNode("fieldSeq"))
		xfName = nullText(xField.selectSingleNode("fieldName"))
		xfLabel = nullText(xField.selectSingleNode("fieldLabel"))
	
%>
<TR><TD><SPAN id="df_<%=xfName%>" title="<%=xfName%>" style="cursor:hand"><%=xfLabel%>
</SPAN></TD></TR>
<% 		
	next
%>
</table>
</body>
<script language=vbs>
	set htPageDom = CreateObject("MICROSOFT.XMLDOM")
	htPageDom.async = false
	LoadXML = "ws/processparamField.asp?fname="

sub document_onClick
	set srcItem = window.event.srcElement
  if srcItem.tagName = "SPAN" then
'		msgBox srcItem.id
		grid = srcItem.id
'		msgBox parent.parent.frames("fedit").document.all.tags("SPAN").length
		set allSpan = parent.parent.frames("fedit").document.all.tags("SPAN")
		for each xs in allSpan
			if xs.id = grid then
				srcItem.outerHTML = ""
				exit sub
			end if
		next
	xStyle=" id="&grid&" style=""background-color:#EEE;"">"
	xv = htPageDom.load(LoadXML&mid(grid,4))
  	if htPageDom.parseError.reason <> "" then 
    		msgBox("htPageDom parseError on line " &  htPageDom.parseError.line)
    		msgBox("Reasonxx: " &  htPageDom.parseError.reason)
    		exit sub
  	end if
	writeCodeStr = nullText(htPageDom.selectSingleNode("//pFieldHTML"))
	xHTML = "<SPAN "&xstyle&writeCodeStr&"</SPAN><BR/>"
	call parent.parent.frames("fedit").insertField (xHTML,grid)
	srcItem.outerHTML = ""


  end if
end sub
</script>
</html>