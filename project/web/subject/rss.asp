<%
'#####(AutoGen) 開始：此段程式碼為自動產生，註解部份請勿刪除 #####
'下列有設定部份如有需要可自行設定
'此版本為  Ver.0.2
'此段程式碼產生日期為： 2009/6/9 上午 10:55:22
'(可修改)未來是否自動更新此程式中的 Pattern (Y/N) : Y

'(可修改)此程式是否記錄 Log 檔
activeLog4U=true

'(可修改)錯誤發生時要到那個頁面,未填時直接結束，1. 可以是 / :首頁, 2. http://xxx.xxx.xxx.xxx/ooo.htm
onErrorPath="/"

'目前程式位置在
progPath="D:\hyweb\GENSITE\project\web\subject\rss.asp"

'--------- (可修改)要檢查的 request 變數名稱 --------
genParamsArray=array("xNode", "count")
genParamsPattern=array("<", ">", "%3C", "%3E", ";", "%27", "'", "=", "+", "-", "*", "/", "%", "@", "~", "!", "#", "$", "&", "\", "(", ")", "{", "}", "[", "]")	'#### 要檢查的 pattern(程式會自動更新):GenPat ####
genParamsMessage=now() & vbTab & "Error(1):REQUEST變數含特殊字元"

'-------- (可修改)只檢查單引號，如 Request 變數未來將要放入資料庫，請一定要設定(防 SQL Injection) --------
sqlInjParamsArray=array()
sqlInjParamsPattern=array("'")	'#### 要檢查的 pattern(程式會自動更新):DB ####
sqlInjParamsMessage=now() & vbTab & "Error(2):REQUEST變數含單引號"

'-------- (可修改)只檢查 HTML標籤符號，如 Request 變數未來將要輸出，請設定 (防 Cross Site Scripting)--------
xssParamsArray=array()
xssParamsPattern=array("<", ">", "%3C", "%3E")	'#### 要檢查的 pattern(程式會自動更新):HTML ####
xssParamsMessage=now() & vbTab & "Error(3):REQUEST變數含HTML標籤符號"

'-------- (可修改)檢查數字格式 --------
chkNumericArray=array()
chkNumericMessage=now() & vbTab & "Error(4):REQUEST變數不為數字格式"

'-------- (可修改)檢察日期格式 --------
chkDateArray=array()
chkDateMessage=now() & vbTab & "Error(5):REQUEST變數不為日期格式"

'##########################################
ChkPattern genParamsArray, genParamsPattern, genParamsMessage
ChkPattern sqlInjParamsArray, sqlInjParamsPattern, sqlInjParamsMessage
ChkPattern xssParamsArray, xssParamsPattern, xssParamsMessage
ChkNumeric chkNumericArray, chkNumericMessage
ChkDate chkDateArray, chkDateMessage
'--------- 檢查 request 變數名稱 --------
Sub ChkPattern(pArray, patern, message)
	for each str in pArray	
		p=request(str)
		for each ptn in patern
			if (Instr(p, ptn) >0) then
				message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
				Log4U(message) '寫入到 log
				OnErrorAction
			end if
		next
	next
End Sub

'-------- 檢查數字格式 --------
Sub ChkNumeric(pArray, message)
	for each str in pArray
		p=request(str)
		if not isNumeric(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'--------檢察日期格式 --------
Sub ChkDate(pArray, message)
	for each str in pArray
		p=request(str)
		if not IsDate(p) then
			message = message & vbTab & progPath & vbTab & "request(" & str & ")=" & p & vbTab & request.serverVariables("REMOTE_ADDR") & vbTab & Request.QueryString
			Log4U(message) '寫入到 log
			OnErrorAction
		end if
	next
End Sub

'onError
Sub OnErrorAction()
	if (onErrorPath<>"") then response.redirect(onErrorPath)
	response.end
End Sub

'Log 放在網站根目錄下的 /Logs，檔名： YYYYMMDD_log4U.txt
Function Log4U(strLog)
	if (activeLog4U) then
		fldr=Server.mapPath("/") & "/Logs"
		filename=Year(Date()) & Right("0"&Month(Date()), 2) & Right("0"&Day(Date()),2)
		
		filename = filename & "_log4U.txt"
		
		Dim fso, f
		Set fso = CreateObject("Scripting.FileSystemObject")
		
		'產生新的目錄
		If (Not fso.FolderExists(fldr)) Then
			Set f = fso.CreateFolder(fldr)
		Else
			ShowAbsolutePath = fso.GetAbsolutePathName(fldr)
		End If
		
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		'開啟檔案
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile( fldr & "\" & filename , ForAppending, True, -1)
		f.Write strLog  & vbCrLf
	end if
End Function
'##### 結束：此段程式碼為自動產生，註解部份請勿刪除 #####
%><?xml version="1.0" encoding="UTF-8" ?> 
<% 
	response.charset = "utf-8"
    Response.Expires = 0
	response.ContentType = "text/xml" 
%>
<!--#Include virtual = "/inc/client.inc" -->
<!--#Include virtual = "/inc/dbFunc.inc" -->
<!--#include virtual = "/inc/web.sqlInjection.inc" -->
 <!--#include virtual = "/inc/checkURL.inc" -->
 <!--#include virtual = "/subject/include/CheckPoint_BeforeLoadXML.asp" -->
<rss version="2.0">
<% 
'this function in include file
call CheckPoint_BeforeLoadXML(request("mp"), request("xnode"), request("xItem"))	


call CheckURL(Request.QueryString)
on error resume next
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open session("ODBCDSN")

 	xTreeNode = 1630
	
	if pkStrWithSriptHTML(request("xNode"), "") <> "null"	then	xTreeNode = pkStrWithSriptHTML(request("xNode"), "")
	
if isnumeric(request("xNode"))  Then
if CInt(request("xNode")) > 0 then
 	
 	set RS = conn.execute("SELECT * FROM CatTreeNode WHERE CtNodeID IN (" & xTreeNode & ")")
  if not RS.eof then
 	xDataNode = RS("CtUnitID")
	HeaderCount = 20
 	if request("count")<>""	then	HeaderCount=request("count")

	sql = "SELECT u.*, u.ibaseDsd, b.sbaseTableName, r.pvXdmp, catName FROM CatTreeNode AS n LEFT JOIN CtUnit AS u ON u.ctUnitId=n.ctUnitId" _
		& " Left Join BaseDsd As b ON u.ibaseDsd=b.ibaseDsd" _
		& " LEFT JOIN CatTreeRoot AS r ON r.ctRootId=n.ctRootId" _
		& " WHERE n.ctNodeId=" & pkstr(xTreeNode,"")
	set xRS = conn.execute(sql)
	
 	response.write "<channel>"
	response.write "<title>農業知識入口網</title>" 
	response.write "<link>http://kmweb.coa.gov.tw</link>"
	response.write "<description>農委會 - " & xRS("CtUnitName") & "</description>" 
	response.write "<language>zh-tw</language>" 
	response.write "<lastBuildDate>" & ServerDatetimeGMT(now()) & "</lastBuildDate>" 
	response.write "<ttl>" & HeaderCount & "</ttl>" 

  		
	if HeaderCount<>"" then HeaderCount = "TOP " & HeaderCount
	
	sql = "SELECT " & HeaderCount & " htx.*, xr1.deptName, u.CtUnitName " _
		& " FROM CuDTGeneric AS htx LEFT JOIN dept AS xr1 ON xr1.deptID=htx.iDept" _
		& " LEFT JOIN CtUnit AS u ON u.CtUnitID=htx.iCtUnit" _
		& " WHERE htx.fCTUPublic='Y' "
	if xDataNode<>"" then		sql = sql & " AND iCtUnit=" & xDataNode
	SqlOrderBy = "xPostDate DESC"
	if SqlOrderBy<>"" then	sql = sql & " ORDER BY " & SqlOrderBy
	set RS = conn.execute(sql)
	while not RS.eof
	    xURL = session("myWWWSite") & "/content.asp?cuItem=" & RS("iCuItem")
    	'xURL = session("mySiteURL") & "/content.asp?cuItem=" & RS("iCuItem") & "&amp;ctNode=" & xTreeNode
    	'xURL = session("mySiteURL") & "/content.asp?xItem=" & RS("iCuItem") & "&amp;ctNode=" & xTreeNode
    	if RS("ibaseDSD") = 2 then	xURL = deAmp(RS("xURL"))
    	if RS("ibaseDSD") = 9 then	xURL = deAmp(RS("xURL"))
    	if RS("showType") = 2 then	xURL = deAmp(RS("xURL"))
		xPostDateStr = ""
		if not isNull(RS("xPostDate")) then xPostDateStr = ServerDatetimeGMT(RS("xPostDate"))    	
%>
		<item iCuItem="<%=RS("iCuItem")%>" newWindow="<%=RS("xNewWindow")%>">
			<title><%=RS("sTitle")%></title>
			<link><%=deAmp(xURL)%></link>
			<guid isPermaLink="false"><%=RS("iCuItem")%></guid>
 			<pubDate><%=xPostDateStr%></pubDate>
			<description><![CDATA[<%=RS("xBody")%>]]></description>
<%		if not isNull(RS("xImgFile")) then %>
			<xImgFile>public/Data/<%=RS("xImgFile")%></xImgFile>
<%		end if %>
		</item>
<%		
		RS.moveNext
	wend
  	response.write "</channel>"
  end if
end if
end if

Function ServerDatetimeGMT(d)             '日期轉換為RSS GMT格式
	DTGMT = dateAdd("h",-8,d)
	WeekDayStr = ""
	Select Case Weekday(DTGMT)
		Case 1 : WeekDayStr = "Sun"
		Case 2 : WeekDayStr = "Mon"
		Case 3 : WeekDayStr = "Tue"
		Case 4 : WeekDayStr = "Wed"
		Case 5 : WeekDayStr = "Thu"
		Case 6 : WeekDayStr = "Fri"
		Case 7 : WeekDayStr = "Sat"
	End Select 
	MonthStr = ""
	Select Case Month(DTGMT)
		Case 1  : MonthStr = "Jan"
		Case 2  : MonthStr = "Feb"
		Case 3  : MonthStr = "Mar"
		Case 4  : MonthStr = "Apr"
		Case 5  : MonthStr = "May"
		Case 6  : MonthStr = "Jun"
		Case 7  : MonthStr = "Jul"
		Case 8  : MonthStr = "Aug"
		Case 9  : MonthStr = "Sep"
		Case 10 : MonthStr = "Oct"
		Case 11 : MonthStr = "Nov"
		Case 12 : MonthStr = "Dec"
	End Select 	
	xhour = right("00" + cstr(hour(DTGMT)),2)
	xminute = right("00" + cstr(minute(DTGMT)),2)
	xsecond = right("00" + cstr(second(DTGMT)),2)
	if Len(d) = 0 then
	  	ServerDatetimeGMT = ""
	else
	 	ServerDatetimeGMT = WeekDayStr + ", " + right("00" + cstr(day(DTGMT)),2) + " " + MonthStr + " " + cStr(Year(DTGMT)) +  " " + xhour + ":" + xminute + ":" + xsecond + " GMT"
	end if
End Function

function nullText(xNode)
  on error resume next
  xstr = ""
  xstr = xNode.text
  nullText = xStr
end function

function deAmp(tempstr)
  xs = tempstr
  if xs="" OR isNull(xs) then
  	deAmp=""
  	exit function
  end if
  	deAmp = replace(xs,"&","&amp;")
end function
%>
</rss>
