<html>
<body style="margin-left:0; margin-top:3;margin-right:0">
<table cellspacing=2 border=1>
<% 

	Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn.Open session("ConnString")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ConnString")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	
	sql = "SELECT p.*, h.posTop FROM INFO_PART AS p LEFT JOIN INFO_HPSET AS h " _
		& " ON p.part_id=h.part_id AND h.user_id='" & session("userid") & "'" _
		& " WHERE p.pubType='系統'"
	set RS = conn.execute(sql)
	
	while not RS.eof
		if isNull(RS("posTop")) then
%>
<TR><TD title="<%=RS("part_desc")%>"><SPAN id=d-<%=RS("part_id")%> style="cursor:hand"><%=RS("part_name")%>
</SPAN></TD></TR>
<% 		
		end if
		RS.moveNext
	wend
%>
</table>
</body>
<script language=vbs>
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
	xStyle=" id="&grid&" style=""border-width:2; border-style:outset; border-color:#999999; background-color:#cccccc" & _
		";position:absolute; top:100; left:10; cursor:hand; width:300px; height:95px; z-index:0;"">"
'	msgBox xStyle
	xImg="<IMG id=vbt" & mid(grid,3) & " src=""bt_updown.gif"" style=""position:absolute; top:185;" _
 		& " left:10;"" title=""拉動調整單元高度"">"
	xHTML = "<SPAN "&xstyle&srcItem.innerText&"</SPAN>" & xImg
	call parent.parent.frames("fedit").document.body.insertAdjacentHTML ("BeforeEnd",xHTML)
	srcItem.outerHTML = ""


	end if
end sub
</script>
</html>
