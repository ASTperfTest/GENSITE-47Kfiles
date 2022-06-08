<%@ CodePage = 65001 %><html>
<head>
<title><%=session("mySiteName")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="MenuFiles/menu_type_learner.css">
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">
<script language=vbs>
dim xmenu(199,4)
dim xmCount(20,1)
dim xtop

xtop = 0

<%
dim menuCat(20)
dim xmenu(199,4)
dim xmCount(20,1)

set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


xn = 0
xmIdx = 1
xItemCount = 0
xPos=Instr(session("uGrpID"),",")
if xPos>0 then
	IDStr=replace(session("uGrpID"),", ","','")
	IDStr="'"&IDStr&"'"
else
	IDStr="'" & session("ugrpID") & "'"
end if
sqlcom = "SELECT DISTINCT AP.APcode, APnameC, APorder, APcatCName, APcat.APseq, APpath " _
		& ", AP.xsNewWindow, AP.xsSubmit" _
		& " FROM AP JOIN APcat ON AP.APCat=APcat.APCatID" _
		& " JOIN uGrpAP ON AP.APCode=uGrpAP.APcode" _
		& " WHERE uGrpID IN (" & IDStr & ") AND rights>0 " _
		& " ORDER BY APcat.APseq, AP.APorder, AP.APCode"	
set RS = Conn.execute(sqlcom)
'response.write sqlCom & vbCRLF

xapcat = ""
xapcode = ""
while not RS.eof

	if RS("APcatCName") <> xapcat then
		
		xn = xn + 1
		menuCat(xn) = RS("APcatCName")
		xmCount(xn-1,1) = xItemCount
		xmCount(xn,0) = xmIdx
%>
		xmCount(<%=xn-1%>,1) = <%=xItemCount%>
		xmCount(<%=xn%>,0) = <%=xmIdx%>
<%
		xapcat = RS("APcatCname")
		xItemCount = 0
		xaporder = ""
	end if

    if RS("APnameC") <> xapcode then	
	xapo = left(RS("APorder"),1)
	if xapo <> xaporder then
		if xitemCount <> 0 then
			xmenu(xmIdx,2) = "Y"
%>
			xmenu(<%=xmIdx%>,2) = "Y"
<%
		end if
		xaporder = xapo
	end if

	xmenu(xmIdx,0) = RS("APnameC")
	xmenu(xmIdx,1) = RS("APpath")
	xmenu(xmIdx,3) = RS("xsNewWindow")
	xmenu(xmIdx,4) = RS("xsSubmit")
%>
	xmenu(<%=xmIdx%>,0) = "<%=RS("APnameC")%>"
	xmenu(<%=xmIdx%>,1) = "<%=RS("APpath")%>"
	xmenu(<%=xmIdx%>,3) = "<%=RS("xsNewWindow")%>"
	xmenu(<%=xmIdx%>,4) = "<%=RS("xsSubmit")%>"
<%
	xmIdx = xmIdx + 1
	xItemCount = xItemCount + 1
	xapcode=RS("APnameC")
    end if
	RS.moveNext
wend
		xmCount(xn,1) = xItemCount
%>
		xmCount(<%=xn%>,1) = <%=xItemCount%>

sub menuClick(xi)
	Dim miid
	Dim mi
	
	xtop = xi
	popMenu.clear
	for i = xmCount(xi,0) to xmCount(xi,0)+xmCount(xi,1)-1
		if xmenu(i,2) = "Y" then
			popmenu.AddItem ""
		end if
		popmenu.AddItem xmenu(i,0)
	next
'	popmenu.popUp 5+(xi-1)*90,87
	miid = "apcat" & Cstr(xi)
	Set mi = document.all(miid)
	popmenu.popUp mi.offsetLeft,87
	Set mi = Nothing
end sub

sub popmenu_Click(item)
'	msgBox item & " " & xtop & " " & xmCount(xtop,0),,xmenu(xmCount(xtop,0)+item-1,1)
'	msgBox xmenu(xmCount(xtop,0)+item-1,0)
'	msgBox xmenu(xmCount(xtop,0)+item-1,1)
'	parent.mainFrame.location = xmenu(xmCount(xtop,0)+item-1,1) & "?menuname=" & xmenu(xmCount(xtop,0)+item-1,0)
  if xmenu(xmCount(xtop,0)+item-1,3)="Y" then
  	window.open xmenu(xmCount(xtop,0)+item-1,1)
  else
	parent.mainFrame.location = xmenu(xmCount(xtop,0)+item-1,1)
  end if
	
end sub

 Sub menuimgon()
'	document.domain = "csmc.edu.tw" 
 	if window.parent.f.cols="0,*" Then
		window.parent.f.cols="159,*"
		menuimg.Src="images/x-2.gif"
	Else
		window.parent.f.cols="0,*"
		menuimg.Src="images/x-1.gif"
	End If
 End Sub
</SCRIPT>
<TABLE id=menu 
style="BACKGROUND: url(/MenuFiles/learner_menu_bg2.gif); LEFT: 0px; POSITION: absolute; TOP: 0px" 
height="100%" cellSpacing=0 cellPadding=0 width=160 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top>
      <TABLE 
      height="100%" cellSpacing=0 cellPadding=0 width=160 border=0>
        <TBODY>
        <TR height=17>
          <TD vAlign=top style="background-color:#3366CC; color:#cccccc; font-weight:bold;" id=xhome><%=session("mySiteName")%></TD></TR>
        <TR>
          <TD vAlign=top>
<%
	for xi = 1 to xn
'		response.write xmCount(xi,0) & "," & xmCount(xi,1) & ")"
%>
        <DIV class=folder onclick="VBS: toggleFolder(me)"><A title="<%=menuCat(xi)%>"
        href="stackMenu.asp#"><IMG 
        height=16 alt="按一下就會收合" src="MenuFiles/folder_close.gif" 
        width=16 border=0><%=menuCat(xi)%></A> </DIV>
        <DIV class=hide>
<%
	for i = xmCount(xi,0) to xmCount(xi,0)+xmCount(xi,1)-1
'		if xmenu(i,2)="Y" then	response.write "<DIV height=1><HR/></DIV>"

		xTarget = "mainFrame"
		if xmenu(i,3)="Y" then	xTarget="_blank"
%>
            <DIV class=item onclick="VBS: takeClick(me)"><A title="<%=xmenu(i,0)%>"
            href="<%=xmenu(i,1)%>" 
            target="<%=xTarget%>"><IMG height=16 
            src="MenuFiles/piece_default.gif" width=16 
            border=0><%=xmenu(i,0)%> </A></DIV>
<%
	next
%>

        </DIV>
<%
	next
%>


        </TD></TR>
</TBODY></TABLE></TD></TR></TBODY></TABLE>
</body>  
</html>  
<script language=javaScript>  
function xmText(msg) {  
	xHelpMsg.innerText = msg  
}  
</script>  
<script language=VBS>
dim openFolder
dim takeItem
set takeItem = xHome

sub CalendarClick
	window.parent.frames(2).navigate "Calendar.asp"
end sub

sub window_onLoad
'	window.top.topFrame.menuimgon()

end sub

sub toggleFolder(xme)
'	set eSrc = window.event.srcElement
'	msgBox eSrc.tagName
'	msgBox xme.tagName
'on error resume next
	if isObject(openFolder) then
		openFolder.className = "folder"
		openFolder.nextSibling.className = "hide"
		openFolder.children(0).children(0).src = "MenuFiles/folder_close.gif"
'		msgBox "hide"
	end if
	
	if openFolder = xme then
'		msgBox "haha"
'		set openFolder = ""
'		exit sub
	end if
	set openFolder = xme
	if xme.className="folder" then
		xme.className = "folderselect"
		xme.nextSibling.className = "show"
		xme.children(0).children(0).src = "MenuFiles/folder_open.gif"
'		msgBox "show"
	else
		xme.className = "folder"
	end if
end sub

sub takeClick(xme)
	if isObject(takeItem) then
		takeItem.className = "item"
'		takeItem.children(0).children(0).src = "MenuFiles/piece_default.gif"
	end if
	set takeItem = xme
	xme.className = "itemselect"
	xme.children(0).children(0).src = "MenuFiles/piece_select.gif"
end sub
</script>
