<%@ CodePage = 65001 %>
<html> 
 
<head> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"> 
<meta name="GENERATOR" content="Microsoft FrontPage 3.0"> 
<style>
.selName { background-color:darkolivegreen; color:white; }
</style>
</head> 
 
<body style="background-image=;">
<form name=regx method="POST" action="saveProfile.aspx">
<INPUT NAME="hpSetStr" type="hidden">
<table border=0>
<tr><td>
<button id=btCheck>儲存</button></td>
<td style="font-size:10pt;">
‧雙擊單元可更換單元底色及高度<br>
‧不要的單元移到橫線上即表刪除，平移橫線上箭頭調整欄寬
</td></table>
</form><HR>
<%
topCount = 0

FUNCTION CkStr (s, endchar)
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	CkStr="'" & s & "'" & endchar
END FUNCTION

	Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	
	
	xhDivPos = 650

	sql = "SELECT * FROM AP " _
		& " WHERE 1=2"
	set RS = conn.execute(sql)
	
	while not RS.eof
%>
<SPAN id=df_<%=RS("part_id")%> title="<%=RS("posHeight")%>" style="align:center; border-width:2; border-style:outset; border-color:#999999; background-color:<%=RS("bgcolor")%>;
position:absolute; top:<%=RS("posTop")%>; left:<%=RS("posLeft")%>; cursor:hand; width:300px; height:<%=RS("posHeight")%>; z-index:0;">
<%=RS("part_name")%></SPAN>
<!--IMG id=vbt<%=RS("part_id")%> src="bt_updown.gif" style="position:absolute; top:<%=RS("posTop")+RS("posHeight")-10%>;
 left:<%=RS("posLeft")%>;" title="拉動調整單元高度"-->
<% 		
		RS.moveNext
	wend
%>
<!--IMG id=hDivSet src="arrow01_10.gif" style="position:absolute; top:70; left:<%=xhDivPos%>;" title="拉動調整欄寬"-->
<script language=vbs>
dim xmid, colorID, cntField, xhDivPos
dim xsaCount
dim xSpanArray(50,5)
dim xSortArray(50,2)

xhDivPos = <%=xhDivPos%>
xmid = ""
colorID = ""
dxmid = ""
xtype = ""

	doSorting

sub btCheck_onClick
'	window.open("/EKP/include/emplist.jsp?id=doc_cat_mod.keeper&name=doc_cat_mod.keeper_name&num=2")
	set allSpan = document.all.tags("SPAN")

	xStr = ""
	for each xs in allSpan
'		msgBox xs.id & "==>" & xs.style.top
	  if xs.style.posTop >= 100 then
		xStr = xStr & xs.id & "," & xs.style.posLeft _
			& "," & xs.style.posTop & "," & xs.style.backgroundColor _
			& "," & xs.style.posHeight & ";" ' & vbCRLF
	  end if
	next
'	msgBox xStr
	regx.hpSetStr.value = xStr
	regx.submit
	
end sub

sub insertField (xhtm, grid)
	window.document.body.insertAdjacentHTML "BeforeEnd",xhtm
	set mObj = document.all(grid)
'	msgBox mObj.clientWidth&","&mObj.clientHeight
		mObj.style.width = mObj.clientWidth+10
		mObj.style.height = mObj.clientHeight
'	msgBox mObj.style.width&","&mObj.style.height
		
	doSorting
'	msgBox xhtm
end sub

sub doSorting
	set allSpan = document.all.tags("SPAN")
	
	xsaCount = 0
	for each xs in allSpan
'		msgBox xs.id & "==>" & xs.style.top
		xsaCount = xsaCount + 1
		xSpanArray(xsaCount,1) = xs.id
		xSpanArray(xsaCount,2) = xs.style.posTop
		xSpanArray(xsaCount,3) = xs.style.posLeft
		xSpanArray(xsaCount,4) = xs.style.posHeight
	next
'	msgBox xsaCount
	sortPos 40
	sortPos xhDivPos
end sub

sub changeDivPos
	set allSpan = document.all.tags("SPAN")
	orgDivPos = xhDivPos
	xhDivPos = hDivSet.style.posLeft
	c1Width = xhDivPos - 50
	
	xsaCount = 0
	for each xs in allSpan
		if xs.style.posLeft = 40 then
			xs.style.Width = c1Width
		else
			xs.style.posLeft = xhDivPos
			xs.style.Width = 600 - c1Width
		end if
	next
	doSorting
end sub

sub sortPos(xLeft)
	xsrCount = 0
	for xi = 1 to xsaCount
		if xSpanArray(xi,3)= xLeft then
			xsrCount = xsrCount + 1
			xSortArray(xsrCount,1) = xi
			xSortArray(xsrCount,2) = xSpanArray(xi,2)
		end if
	next
'	msgBox xsrCount
	rxTop = 100
	c1Width = xhDivPos - 50
	if xLeft <> 40 then		c1Width = 600 - c1Width
	for xi = 1 to xsrCount
		xTopV = 900
		xTopI = 0
		for xj = 1 to xsrCount
			if xSortArray(xj,2) > 50 AND xSortArray(xj,2)<xTopV then 
				xTopI = xj
				xTopV = xSortArray(xj,2)
			end if
		next
'		msgBox xSpanArray(xSortArray(xTopI,1),1)
		cxID = mid(xSpanArray(xSortArray(xTopI,1),1),4)
		document.all("df_"&cxID).style.top = rxTop
		document.all("df_"&cxID).style.width = c1Width
		document.all("df_"&cxID).title = xSpanArray(xSortArray(xTopI,1),4)
'		document.all("vbt"&cxID).style.top = rxTop + document.all("df_"&cxID).style.posHeight - 10
'		document.all("vbt"&cxID).style.left = xLeft
		rxTop = rxTop +  xSpanArray(xSortArray(xTopI,1),4) + 5
		xSortArray(xTopI,2) = 0
	next
end sub

function checkParent(src, dest) 
on error resume next
	while not src.tagName="BODY"
		if src.tagName = dest then 
			set checkParent = src
			exit function
		end if
		set src = src.parentElement
	wend
	set checkParent = src
end function

sub document_onDblClick
  set mObj = window.event.srcElement
'  msgBox mObj.id
  if mObj.tagName="SPAN" then
  	fcolor=showModalDialog("editor_color.htm",false,"dialogWidth:106px;dialogHeight:156px;status:0;")
	if fcolor="" then	exit sub
'  	msgBox fcolor
	xp = split(fcolor,";")
  	mObj.style.backgroundColor = xp(0)
  	nHeight = xp(1)
  	if isNumeric(nHeight) then
  		if nHeight > 50 AND nHeight < 500 then
  			mObj.style.height = nHeight & "px"
  			doSorting
  		end if
  	end if
  end if 
    window.event.cancelBubble = true
    window.event.returnValue = false
end sub

sub document_onMouseDown
  set mObj = window.event.srcElement
'	msgBox mObj.tagName
  set xObj = mObj
  set pObj = checkParent(xObj,"SPAN")
'	msgBox xObj.tagName
'	msgBox mObj.tagName
	if (xmid = "") and (mObj.tagName = "IMG") then 
		xtype = "IMG"
		xmid = mObj.id
'		mObj.style.borderStyle = "groove"
		mObj.style.zIndex = 1
		window.status = "調整寬度"
	elseif (xmid = "") and (mObj.tagName = "SPAN") then 
		set mObj = pObj
		xtype = "SPAN"
		xmid = mObj.id
		dxmid = mObj.id
		mObj.style.borderStyle = "groove"
		mObj.style.zIndex = 1
		mObj.style.height = mObj.clientHeight
		mObj.style.width = mObj.clientWidth+20
		window.status = "搬移欄位" & mObj.clientHeight
	else
		if xmid<>"" then	document.all(xmid).style.borderStyle = "none"
		xmid = ""
	end if
end sub

sub document_onMouseMove
	if xmid <> "" then
		set mObj = document.all(xmid)
'		mObj.style.top = (window.event.y)-10-mObj.Height\2
'		mObj.style.left = (window.event.x)-20-mObj.Width\2
		if xmid = "hDivSet" then
			mObj.style.left = (window.event.x)-10
			window.status = "設定左欄寬度: " & (mObj.style.posleft-50) & " px"
		elseif left(xmid,3)="vbt" then
			window.status = mObj.style.top & "設定單元高度" & window.event.screenY
			mObj.style.top = (window.event.Y+document.body.scrollTop)-5
'			window.status = "設定單元高度"
		else
			mObj.style.left = (window.event.x)-10
			mObj.style.top = (window.event.y+document.body.scrollTop)-10
		end if
	else
'			window.status = window.event.Y & "-" & document.body.scrollTop
	end if
  ' prevent event being handled elsewhere and the default action
  window.event.cancelBubble = true
  window.event.returnValue = false
end sub

sub document_onMouseUp
  if xmid <> "" then
'  	msgBox xmid

	set xObj = document.all(xmid)
'  	msgBox xObj.style.left
	if xmid="hDivSet" then
		if xObj.style.posLeft < 200 then
	  		xObj.style.left = 200
	  	elseif xObj.style.posLeft > 600 then
	  		xObj.style.left = 600
	  	end if
	elseif left(xmid,3)="vbt" then
		set spxObj=document.all("df_"&mid(xmid,4))
'		msgbox spxObj.style.posTop & "," & xObj.style.posTop
		newHeight = xObj.style.posTop - spxObj.style.posTop
'		msgbox newHeight
		if newHeight >= 50 then	
			spxObj.style.posHeight = newHeight
		else
			xObj.style.posTop = spxObj.style.posTop + spxObj.style.posHeight - 10
		end if
	else
	  if xObj.style.posTop > 50 then
		if xObj.style.posLeft < xhDivPos then
			xObj.style.left = 40
		else
			xObj.style.left = xhDivPos
		end if
	  end if
	end if

    ' return the piece to the lower z-order position
    document.all(xmid).style.zIndex = 0
	'document.all(xmid).style.borderStyle = "none"

    ' reset the global 'dragging' variable
  
    ' prevent event being handled elsewhere and the default action
    window.event.cancelBubble = true
    window.event.returnValue = false
    
	if xmid<>"hDivSet" then
	    doSorting
	else
    	changeDivPos
    end if

    xmid = ""
    window.status = "拖拉任何欄位名稱可移動位置"

  end if
end sub

</script>
</body> 
</html> 
