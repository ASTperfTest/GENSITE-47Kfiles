<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="點閱統計"
HTProgFunc="點閱統計"
HTProgCode="GW1M22"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%

'取得某節點的點閱次數
Function GetAccessCount(nodeID)

	Set rs77 = Server.CreateObject("ADODB.RecordSet")
	sql77 = " select count(*) from gipHitUnit where hitTime between N'"& date1 &"' and N'"& date2 &"' and iCtNode=N'"& nodeID &"' "

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rs77.Open sql77, conn, 1, 3
Set rs77 =  conn.execute(sql77)

'----------HyWeb GIP DB CONNECTION PATCH----------

	cnt = rs77(0)
	rs77.close
	set rs77=nothing
	GetAccessCount = cnt
	
End Function


'遞迴處理每個子節點
Sub ProcessChildNode(rootID, parentID)
	Set rs95 = Server.CreateObject("ADODB.RecordSet")
	
	sql95 = " select * from CatTreeNode Where CtRootID=N'"& rootID &"' and DataParent = N'"& parentID &"' order by CatShowOrder "
	'response.write sql95 &"<BR>"

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rs95.Open sql95, conn, 1, 3
Set rs95 =  conn.execute(sql95)

'----------HyWeb GIP DB CONNECTION PATCH----------

	if not rs95.EOF then
%>
	<ul>
<%	
		while not rs95.EOF
			accessCNT = GetAccessCount(trim(rs95("CtNodeID")))
			response.write "<li>"& trim(rs95("CtNodeID")) &" - "& trim(rs95("CatName")) &" - <B>"& accessCNT &"</B></li>"
						
			ProcessChildNode rootID, trim(rs95("CtNodeID"))
				
			'if trim(rs95("CtNodeKind"))="C" then
			'	ProcessChildNode rootID, trim(rs95("CtNodeID"))
			'else
			'	accessCNT = GetAccessCount(trim(rs95("CtNodeID")))
			'	response.write "<B>"& accessCNT &"</B>"
			'end if		
			
			rs95.movenext
		wend
%>
	</ul>
<%		
	end if
	rs95.close
	set rs95=nothing
End Sub
 
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>點閱統計</title>
</head>

<body>

<script language=vbs>
Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=document.all.CalendarTarget.value
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=document.all.CalendarTarget.value
        end if
	end if
end sub   

Sub fmQuerySubmit()	
    
	nMsg = "請務必填寫「{0}」，不得為空白！"
	lMsg = "「{0}」欄位長度最多為{1}！"
	dMsg = "「{0}」欄位應為 yyyy/mm/dd ！"
	iMsg = "「{0}」欄位應為數值！"
	pMsg = "「{0}」欄位圖檔類型必須為 " &chr(34) &"GIF,JPG,JPEG" &chr(34) &" 其中一種！"

	flag_1 = false
	for each xri in stat1.all("TreeRoot")
		if xri.checked then 
			flag_1 = true
		end if
	next    
    if not flag_1 then
		MsgBox replace(nMsg,"{0}","目錄樹"), 64, "Sorry!"
		exit sub
    end if
    		
	IF stat1.htx_dbDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料範圍起日"), 64, "Sorry!"
		'stat1.htx_dbDate.focus
		exit sub
	END IF
	IF (stat1.htx_dbDate.value <> "") AND (NOT isDate(stat1.htx_dbDate.value)) Then
		MsgBox replace(dMsg,"{0}","資料範圍起日"), 64, "Sorry!"
		'stat1.htx_dbDate.focus
		exit sub
	END IF
	IF stat1.htx_deDate.value = Empty Then 
		MsgBox replace(nMsg,"{0}","資料範圍迄日"), 64, "Sorry!"
		'stat1.htx_deDate.focus
		exit sub
	END IF
	IF (stat1.htx_deDate.value <> "") AND (NOT isDate(stat1.htx_deDate.value)) Then
		MsgBox replace(dMsg,"{0}","資料範圍迄日"), 64, "Sorry!"
		'stat1.htx_deDate.focus
		exit sub
	END IF
  
  stat1.Submit
End Sub
</script>

<p>網站類別點閱統計</p>
<form name="stat1" action="type_stat.asp" method="post">
<object data="../inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>

目錄樹：<BR>
<%
	Set rs88 = Server.CreateObject("ADODB.RecordSet")
	TreeCnt = 0
	SQL = " select * from CatTreeRoot Where pvXdmp is not null order by CtRootID "

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rs88.Open SQL , conn , 1 , 3
Set rs88 =  conn .execute(SQL )

'----------HyWeb GIP DB CONNECTION PATCH----------

	if not rs88.EOF then
		while not rs88.EOF
			TreeCnt = TreeCnt + 1
			RootID = trim(rs88("CtRootID"))
			RootName = trim(rs88("CtRootName"))
%>
			<input name="TreeRoot" type="checkbox" value="'<%=RootID%>'"><%=RootName%><BR>
<%			
			rs88.movenext
		wend
	else
		response.write "No DATA"
	end if
	rs88.close
	Set rs88=nothing		
%>

資料範圍起日：<input name="htx_dbDate" type=hidden>
<input name="pcShowhtx_dbDate" size="8" readonly onclick="VBS: popCalendar 'htx_dbDate' ,''">

資料範圍迄日：<input name="htx_deDate" type=hidden>
<input name="pcShowhtx_deDate" size="8" readonly onclick="VBS: popCalendar 'htx_deDate' ,''">

<input type="button" value="開始統計" name="B3" onclick="fmQuerySubmit()"></p>
</form>

<%
	date1 = request("htx_dbDate")
	date2 = request("htx_deDate")

if date1<>"" and date2<>"" then

	TreeRoot = request("TreeRoot")

	Set rs = Server.CreateObject("ADODB.RecordSet")
   
	SQL = " select * from CatTreeRoot Where CtRootID in ( "& TreeRoot &" ) "
	

'----------HyWeb GIP DB CONNECTION PATCH----------
'	rs.Open SQL , conn , 1 , 3
Set rs =  conn .execute(SQL )

'----------HyWeb GIP DB CONNECTION PATCH----------

	if not rs.EOF then
		while not rs.EOF

			response.write trim(rs("CtRootID")) &" - "& trim(rs("CtRootName")) &" - "& trim(rs("vGroup")) &"<BR>"
			
			ProcessChildNode trim(rs("CtRootID")),0
		
			rs.movenext
		wend
	else
		response.write "No DATA"
	end if
	
	rs.close
	Set rs=nothing	
	
%>
<p><input type="button" value="轉成Excel" name="B3" onclick="javascript:alert('沒有作用啦');"></p>
<%
end if
%>
</body>

</html>
