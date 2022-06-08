<%
function ServerDate(dt)             'SQL Server 端的系統日期，供Compare時使用
   if Len(dt)=0 Or isnull(dt) then
       ServerDate=""
   else
       ServerDate = Year(now()) & "/" & Month(now()) & "/" & Day(now())
   end if
end function

function ServerDatetime(d)             'SQL Server 端的系統日期，供Compare時使用
        xyear =cstr(year(now))
        xmonth = right("00"+ cstr(month(now)),2)
        xday = right("00"+ cstr(day(now)),2)
        xhour = right("00" + cstr(hour(now)),2)
        xminute = right("00" + cstr(minute(now)),2)
        xsecond=second(now)
      if Len(d)=0 Or isnull(d) then
          ServerDatetime=""
      else
         ServerDatetime = xyear&"/"&xmonth &"/"&xday&" "&xhour&":"&xminute&":"&xsecond
      end if
end function



function xStdTime(dt)                   '轉成民國年及時間  999/99/99 10:00 給資料型態為DateTime 使用
   if Len(dt)=0 or isnull(dt) then
     	xStdTime=""
   else
        xyear = right("000"+ cstr((year(dt)-1911)),3)     '補零
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xhour = right("00" + cstr(hour(dt)),2)
        xminute = right("00" + cstr(minute(dt)),2)
        xStdTime = xyear & "/" & xmonth & "/" & xday & " " & xhour & ":" & xminute
   end if
end function

function xStdDay(dt)                   '轉成民國年及時間  999/99/99 10:00 給資料型態為DateTime 使用
   if Len(dt)=0 or isnull(dt) then
     	xStdDay=""
   else
        xyear = right("0000"+ cstr(year(dt)),4)     '補零
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xStdDay = xyear & "/" & xmonth & "/" & xday 
   end if
end function

function d7date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
	     d7date=""
	else
                      xy=right("000"+ cstr((year(dt)-1911)),2)     '補零
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      d7date=xy & "." & xm & "." & xd
                  end if
end function

Function formatDate(dt)	'補零
	xyear = year(dt)
	xmonth = right("00"+ cstr(month(dt)),2)
	xday = right("00"+ cstr(day(dt)),2)
	formatDate = xyear & "/" & xmonth & "/" & xday
End Function

										'轉成月日年 99/99/9999
function xMonStdDay(dt)
	if Len(dt)=0 or isnull(dt) then
		xMonStdDay = ""
	else
		xyear = right("0000"+ cstr(year(dt)),4)
		xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
		xMonStdDay = xmonth & "/" & xday & "/" & xyear
	end if
end function

function ckSelect(xs1, xs2)
  if xs1=xs2 then
	ckSelect = " selected"
  else
	ckSelect = ""
  end if
end function

FUNCTION CkStr (s, endchar)
  if s="" then
	ckStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	CkStr="N'" & s & "'" & endchar
  end if
END FUNCTION

FUNCTION pkStr (s, endchar)
  if s="" then
	pkStr = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkStr="N'" & s & "'" & endchar
  end if
END FUNCTION

FUNCTION pkDate (s, endchar)
  if s="" then
	pkDate = "null" & endchar
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	pkDate="'" & s & "'" & endchar
  end if
END FUNCTION

FUNCTION dfs (s)
  if s="" then
	dfs = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	dfs="'" & s & "',"
  end if
END FUNCTION

FUNCTION dfd (s)
  if s="" then
	dfd = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	dfd="'" & s & "',"
  end if
END FUNCTION

FUNCTION dfn (s)
  if isNumeric(s) then
	dfn= s & ","
  else
	dfn = "null,"
  end if
END FUNCTION

FUNCTION drs (xs)
  s = request(xs)
  if s="" then
	drs = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	drs="'" & s & "',"
  end if
END FUNCTION

FUNCTION drd (xs)
  s = request(xs)
  if s="" then
	drd = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	drd="'" & s & "',"
  end if
END FUNCTION

FUNCTION drn (xs)
s = request(xs)
  if s="" then
	drn = "null,"
  else
	drn= s & ","
  end if
END FUNCTION

function tagChg(s)                        '將< > 取代  2000/12/14
   while instr(s,"<")>0
         xs=instr(s,"<")
         if xs>0 then
            s=left(s,xs-1) + "&lt" + mid(s,xs+1)
         end if
         xs=instr(s,">")
         if xs>0 then
            s=left(s,xs-1) + "&gt" + mid(s,xs+1)
         end if
   wend
   tagChg=s
end function


Function stripHTML(strHTML)
'Strips the HTML tags from strHTML

 		Dim objRegExp, strOutput
 		Set objRegExp = New Regexp

 		objRegExp.IgnoreCase = True
 		objRegExp.Global = True
 		objRegExp.Pattern = "(<.+?>)|(%3C.+?%3E)|(<.+?%3E)|(%3C.+?>)"

 		'Replace all HTML tag matches with the empty string
 		strOutput = objRegExp.Replace(strHTML, "")
 
 		'Replace all < and > with &lt; and &gt;
 		strOutput = Replace(strOutput, "<", "&lt;")
 		strOutput = Replace(strOutput, ">", "&gt;")
 		strOutput = Replace(strOutput, "%3C", "&lt;")
 		strOutput = Replace(strOutput, "%3E", "&gt;")
 		stripHTML = strOutput 'Return the value of strOutput
 		
 		Set objRegExp = Nothing

End Function

Function pkStrWithSriptHTML(s, endchar)
	pkStrWithSriptHTML=pkStr(stripHTML(s), endchar)
End Function



%>

