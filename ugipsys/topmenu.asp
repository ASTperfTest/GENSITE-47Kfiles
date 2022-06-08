<%@ CodePage = 65001 %><!-- #INCLUDE FILE="inc/dbFunc.inc" -->
<html>
<head>
<title>財政部</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="inc/setstyle.css">
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">
<script language=vbs>
dim xmenu(99,4)
dim xmCount(10,1)
dim xtop

xtop = 0

<%
dim menuCat(10)

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
xFileUpPath = ""
xDataUpPath = ""
xPos=Instr(session("uGrpID"),",")
if xPos>0 then
	IDStr=replace(session("uGrpID"),", ","','")
	IDStr="'"&IDStr&"'"
else
	IDStr="'" & session("ugrpID") & "'"
end if
sqlcom = "SELECT DISTINCT AP.APcode, APnameC, APorder, APcatCName, APcat.APseq, SpecPath AS APpath " _
		& ", AP.xsNewWindow, AP.xsSubmit" _
		& " FROM AP JOIN APcat ON AP.APCat=APcat.APCatID" _
		& " JOIN uGrpAP ON AP.APCode=uGrpAP.APcode" _
		& " WHERE uGrpID IN (" & IDStr & ") AND rights>0 " _
		& " ORDER BY APcat.APseq, AP.APorder, AP.APCode"	
set RS = Conn.execute(sqlcom)

xapcat = ""
xapcode = ""
while not RS.eof

	if RS("APcatCName") <> xapcat then
		
		xn = xn + 1
		menuCat(xn) = RS("APcatCName")
%>
		xmCount(<%=xn-1%>,1) = <%=xItemCount%>
		xmCount(<%=xn%>,0) = <%=xmIdx%>
<%
		xapcat = RS("APcatCname")
		xItemCount = 0
		xaporder = ""
	end if

    if RS("APnameC") <> xapcode then	
    	if RS("APcode")="GC1AP5" then	xFileUpPath = RS("APpath")
    	if RS("APcode")="GC1AP1" then	xDataUpPath = RS("APpath")

	xapo = left(RS("APorder"),1)
	if xapo <> xaporder then
		if xitemCount <> 0 then
%>
			xmenu(<%=xmIdx%>,2) = "Y"
<%
		end if
		xaporder = xapo
	end if

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
	popmenu.popUp mi.offsetLeft,27
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
<%
	sqlCom = "SELECT i.*, d.deptName, t.mValue AS topDataCat " _
		& " FROM InfoUser AS i LEFT JOIN dept AS d ON i.deptID=d.deptID " _
		& " LEFT JOIN CodeMain AS t ON t.codeMetaID='topDataCat' AND t.mCode=i.tDataCat" _
		& " WHERE userID='" & session("userID") & "'"
	set RS = conn.execute(sqlCom)
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#6097CF">
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#6097CF">
        <tr valign="bottom"> 
          <td height="20"> 
            <table border="0" cellspacing="0" cellpadding="0" bgcolor="#6097CF">
	        <tr>
	          <td width="20" align="center"><img ID=menuimg src="images/X-2.gif" alt="收放視窗" width="13" height="13" style="cursor: hand;" onclick="VBScript: menuimgon"></td>
<% 	if Instr(session("uGrpID")&",", "HTSD,") > 0 then %>
<%		for i = 1 to xn %>               
                  <td width="87" align="center" class="tab-title" id="apcat<%=i%>" onMouseOver="Vbscript: document.all.apcat<%=i%>.classname='tab-titleon'"  onMouseOut="Vbscript: document.all.apcat<%=i%>.classname='tab-title'" height="17" onClick="menuClick(<%=i%>)" valign="bottom"><%=menuCat(i)%>               </td>
<%		next
	elseif xFileUpPath<>"" then
		if xDataUpPath<>"" then
%>
     <td width="87" align="center" class="tab-title" id="apxData" onMouseOver="Vbscript: document.all.apxData.classname='tab-titleon'"  onMouseOut="Vbscript: document.all.apxData.classname='tab-title'" height="17" 
     onClick="parent.mainFrame.location= '<%=xDataUpPath%>'" valign="bottom">資料上稿</td>
<%		end if %>
     <td width="87" align="center" class="tab-title" id="apxFile" onMouseOver="Vbscript: document.all.apxFile.classname='tab-titleon'"  onMouseOut="Vbscript: document.all.apxFile.classname='tab-title'" height="17" 
     onClick="parent.mainFrame.location= '<%=xFileUpPath%>'" valign="bottom">檔案上傳</td>
<%	end if %>
		</tr> 
	    </table>
	  </td>	    
	  <td align="center" bgcolor="#6097CF" style="color:#cccccc; font-size:10pt;">
		
	  <%=RS("userName")%> (<%=RS("userID")%>,<%=RS("deptName")%>-<%=RS("topDataCat")%>) 
	  上次於 <%=d7date(session("lastVisit"))%> 由<%=session("lastIP")%>第<%=session("VisitCount")%>次登入
	  </td>           
          <td align="right" valign="bottom" width="10%" height=20>
          <img id=xx src="images/home.gif" width="19" height="20" hspace="3" border="0" alt="回首頁" style="cursor: hand; " onclick="VBS:window.top.mainFrame.location='../main.htm'">
          <img src="images/logout.gif" width="23" height="19" hspace="3" alt="登出" style="cursor: hand; " onclick="VBS:window.navigate 'logout.asp'">  
          </td>  
        </tr>  
      </table>  
    </td>  
  </tr>  
  <tr>   
    <td bgcolor="#052C77"><img src="images/shim.gif" width="5" height="5"></td>  
  </tr>  
</table>  
<% 	if Instr(session("uGrpID")&",", "HTSD,") > 0 then %>
<OBJECT      
	  id=PopMenu      
	  classid="clsid:7823A620-9DD9-11CF-A662-00AA00C066D2"      
	  codebase="iemenu.cab#Version=4,70,0,1161"      
	  align=left      
	  hspace=0      
	  vspace=0 width="14" height="14"      
   >      
</OBJECT>  
<%	end if %>    
</body>  
</html>  
<script language=javaScript>  
function xmText(msg) {  
	xHelpMsg.innerText = msg  
}  
</script>  
<script language=VBS>
sub CalendarClick
	window.parent.frames(2).navigate "Calendar.asp"
end sub

sub window_onLoad
'	window.top.topFrame.menuimgon()

end sub
</script>
