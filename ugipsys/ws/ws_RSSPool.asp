<%
'----941124RSSPool產生機制
cTabchar = chr(9)
function ncTabChar(n)
	if cTabchar = "" then
		ncTabChar = ""
	else
		ncTabChar = String(n,cTabchar)
	end if
end function
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

Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------

SQLC = "Select sTitle,xBody,xPostDate,xImgFile from CuDTGeneric where iCuItem=" & session("RSS_iCuItem")
Set RSC = conn.execute(SQLC)
'SQL = "Select iCuItem from RSSPool where iCuItem=" & session("RSS_iCuItem")
'Set RS = conn.execute(SQL)
if session("RSS_method") = "insert" or session("RSS_method") = "update" then			'----新增insert/編修update
	if not RSC.eof then
		xPostDateStr = ""
		if not isNull(RSC("xPostDate")) then xPostDateStr = ServerDatetimeGMT(now())
		RSSStr = "<item>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<title><![CDATA["&RSC("sTitle")&"]]></title>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<link>"&session("myWWWSiteURL")&"/content.asp?mp="&session("pvXdmp")&"&amp;cuItem="&session("RSS_iCuItem")&"</link>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<guid isPermaLink=""false"" method="""&session("RSS_method")&""">"&session("RSS_iCuItem")&"</guid>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<pubDate>"&xPostDateStr&"</pubDate>" & vbcrlf
	'	RSSStr = RSSStr & ncTabchar(1) & "<category>"&RSC("category")&"</category>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<description><![CDATA["&RSC("xBody")&"]]></description>" & vbcrlf	
		'if not isNull(RSC("xImgFile")) then
	    '        RSSStr = RSSStr & ncTabchar(1) & "<media:content type=""image/jpeg"" url=""" &session("myWWWSiteURL")&"/public/Data/" & RSC("xImgFile") & """/>" & vbcrlf	
		'end if	
		RSSStr = RSSStr & "</item>"
		SQLCreate = "Insert Into RSSPool values("&session("RSS_iCuItem")&",'"&session("RSS_method")&"','"&RSSStr&"',getdate(),'"&session("userID")&"',"&session("ctNodeId")&")"
		conn.execute(SQLCreate)	
	end if
elseif session("RSS_method") = "delete" then		'----刪除delete
	if not RSC.eof then
		RSSStr = "<item>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<title><![CDATA["&RSC("sTitle")&"]]></title>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<link></link>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<guid isPermaLink=""false"" method="""&session("RSS_method")&""">"&session("RSS_iCuItem")&"</guid>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<pubDate>"&RSC("xPostDate")&"</pubDate>" & vbcrlf
	'	RSSStr = RSSStr & ncTabchar(1) & "<category></category>" & vbcrlf
		RSSStr = RSSStr & ncTabchar(1) & "<description><![CDATA["&RSC("xBody")&"]]></description>" & vbcrlf	
		'if not isNull(RSC("xImgFile")) then
        '    RSSStr = RSSStr & ncTabchar(1) & "<xImgFile>"&session("myWWWSiteURL")&"/public/Data/" & RSC("xImgFile") & "</xImgFile>" & vbcrlf	
		'end if	
		RSSStr = RSSStr & "</item>"
		SQLCreate = "Insert Into RSSPool values("&session("RSS_iCuItem")&",'"&session("RSS_method")&"','"&RSSStr&"',getdate(),'"&session("userID")&"',"&session("ctNodeId")&")"
		conn.execute(SQLCreate)	
	end if
end if


%>
