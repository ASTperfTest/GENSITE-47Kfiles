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
sqlcom = "SELECT DISTINCT Ap.apcode, apnameC, aporder, apcatCname, Apcat.apseq, appath " _
		& ", Ap.xsNewWindow, Ap.xsSubmit" _
		& " FROM Ap JOIN Apcat ON Ap.apcat=Apcat.apcatId" _
		& " JOIN UgrpAp ON Ap.apcode=UgrpAp.apcode" _
		& " WHERE ugrpId IN (" & IDStr & ") AND rights>0 " _
		& " AND appath IS NOT NULL" _
		& " ORDER BY Apcat.apseq, Ap.aporder, Ap.apcode"	
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
<script type="text/javascript">
        function setHref(url,target) {
            var isIE = (navigator.appName.indexOf("Microsoft") != -1) ? true : false;
            if (!isIE) {
                parent.location.href = url;
            } else {
                var lha = document.getElementById('_lha');
                if (!lha) {
                    lha = document.createElement('a');
                    lha.id = '_lha';
                    lha.target = target;  // Set target: for IE
                    document.body.appendChild(lha);
                }
                lha.href = url;
                lha.click();
            }
        }
    </script>
</head>
<body style="overflow-x: hidden;overflow-y:auto">
<div id="xFunction" style="display:block;">
<%
	for xi = 1 to xn
'		response.write xmCount(xi,0) & "," & xmCount(xi,1) & ")"
%>
		<a title="<%=menuCat(xi)%>" href="#" class="Cat1" onClick="setMenu(<%=xi%>)"><%=menuCat(xi)%></a>
	<div class="Cat2" id="xGroup<%=xi%>" style="display:none">
<%
	for i = xmCount(xi,0) to xmCount(xi,0)+xmCount(xi,1)-1
		xTarget = "mainFrame"
		if xmenu(i,3)="Y" then	xTarget="_blank"		
		
		if left(xmenu(i,1),10) = "javascript" then
%>  
            <a title="<%=xmenu(i,0)%>" href="<%=xmenu(i,1)%>" ><%=xmenu(i,0)%> </a>
<%
		else	
%>
            <A title="<%=xmenu(i,0)%>"
            href="<%=xmenu(i,1)%>" 
            target="<%=xTarget%>" onclick="setHref('<%=xmenu(i,1)%>','<%=xTarget%>')"><%=xmenu(i,0)%> </A>
<%
        end if
	next
%>
        </DIV>
<%
	next
%>
    </div>
<div class="Display">
<!-- 開啟的狀態, click關閉 

  <a id="switchbotton" href="javascript:;" class="Hide"></a>
-->
</div>
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

sub switchbotton_onClick
	if switchbotton.className = "Hide" then
		xFunction.style.display = "none"
		window.parent.f.cols = "7,*"
		switchbotton.className = "Show"
	else
		xFunction.style.display = "block"
		window.parent.f.cols = "167,*"
		switchbotton.className = "Hide"
	end if

end sub

sub bOpen_onClick
	xMenu.style.display = "block"
	window.parent.f.cols="170,*"
end sub

sub window_onLoad
'	window.top.topFrame.menuimgon()

end sub
</script>
