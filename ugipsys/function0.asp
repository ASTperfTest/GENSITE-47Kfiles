<%@ CodePage = 65001 %><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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

		xapcat = RS("APcatCname")
		xItemCount = 0
		xaporder = ""
	end if

    if RS("APnameC") <> xapcode then	
	xapo = left(RS("APorder"),1)
	if xapo <> xaporder then
		if xitemCount <> 0 then
			xmenu(xmIdx,2) = "Y"
		end if
		xaporder = xapo
	end if

	xmenu(xmIdx,0) = RS("APnameC")
	xmenu(xmIdx,1) = RS("APpath")
	xmenu(xmIdx,3) = RS("xsNewWindow")
	xmenu(xmIdx,4) = RS("xsSubmit")

	xmIdx = xmIdx + 1
	xItemCount = xItemCount + 1
	xapcode=RS("APnameC")
    end if
	RS.moveNext
wend
		xmCount(xn,1) = xItemCount
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/function.css" rel="stylesheet" type="text/css">
</head>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>

<body onload="MM_preloadImages('xslgip/intra1/images/switch_close_2.gif','xslgip/intra1/images/switch_open_2.gif')">
<table width="100%" height="100%" border="0" align="left" cellpadding="0" cellspacing="0" >
  <tr>
    <td id="xMenu" valign="top" style="display:block;">
<div id="Function">
<%
	for xi = 1 to xn
'		response.write xmCount(xi,0) & "," & xmCount(xi,1) & ")"
%>
		<a title="<%=menuCat(xi)%>" href="#" class="Cat1" onClick="setMenu(<%=xi%>)"><%=menuCat(xi)%></a>
		<span id="xGroup<%=xi%>" style="display:none">
	<div id="Cat2">
<%
	for i = xmCount(xi,0) to xmCount(xi,0)+xmCount(xi,1)-1
		xTarget = "mainFrame"
		if xmenu(i,3)="Y" then	xTarget="_blank"
%>
            <A title="<%=xmenu(i,0)%>"
            href="<%=xmenu(i,1)%>" 
            target="<%=xTarget%>"><%=xmenu(i,0)%> </A>
<%
	next
%>
        </DIV></span>
<%
	next
%>
    </div>
  </td>
    <td width="10" height="100%" align="right" valign="middle">
    <a href="#close" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('switchbotton','','xslgip/intra1/images/switch_close_2.gif',1)">
    <img id="bClose" src="xslgip/intra1/images/switch_close_1.gif" alt="縮小選單區" name="switchbotton" width="10" height="52" border="0" id="switchbotton" /></a>
    <a href="#open" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('switchbutton2','','xslgip/intra1/images/switch_open_2.gif',1)">
    <img id="bOpen" src="xslgip/intra1/images/switch_open_1.gif" alt="展開選單區" name="switchbutton2" width="10" height="52" border="0" id="switchbutton2" /></a>
    </td>
  </tr>
</table>
</body>
</html>
<script language=VBS>
	dim initGroup
	initGroup = 0
	setMenu (1)

	
sub setMenu(xn)
	if initGroup <> 0 then	document.all("xGroup"&initGroup).style.display = "none"
	initGroup = xn
	document.all("xGroup"&initGroup).style.display = "block"
end sub
	
sub xxx	
	set xObj = document.all("xGroup"&xn)
	if xObj.style.display = "none" then
		xObj.style.display = "block"
	else
		xObj.style.display = "none"
	end if
end sub
sub bClose_onClick
	xMenu.style.display = "none"
	window.parent.f.cols="10,*"

end sub

sub bOpen_onClick
	xMenu.style.display = "block"
	window.parent.f.cols="170,*"
end sub

sub window_onLoad
'	window.top.topFrame.menuimgon()

end sub
</script>
