<%
function ServerDate(dt)             'SQL Server ºÝªº¨t²Î¤é´Á¡A¨ÑCompare®É¨Ï¥Î
   if Len(dt)=0 Or isnull(dt) then
       ServerDate=""
   else
       ServerDate = Year(now()) & "/" & Month(now()) & "/" & Day(now())
   end if
end function

function ServerDatetime(d)             'SQL Server ºÝªº¨t²Î¤é´Á¡A¨ÑCompare®É¨Ï¥Î
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



function xStdTime(dt)                   'Âà¦¨¥Á°ê¦~¤Î®É¶¡  999/99/99 10:00 µ¹¸ê®Æ«¬ºA¬°DateTime ¨Ï¥Î
   if Len(dt)=0 or isnull(dt) then
     	xStdTime=""
   else
        xyear = right("000"+ cstr((year(dt)-1911)),3)     'è£œé›¶
        xmonth = right("00"+ cstr(month(dt)),2)
        xday = right("00"+ cstr(day(dt)),2)
        xhour = right("00" + cstr(hour(dt)),2)
        xminute = right("00" + cstr(minute(dt)),2)
        xStdTime = xyear & "/" & xmonth & "/" & xday & " " & xhour & ":" & xminute
   end if
end function

function d7date(dt)     'Âà¦¨¥Á°ê¦~  999/99/99 µ¹¸ê®Æ«¬ºA¬°SmallDateTime ¨Ï¥Î
	if Len(dt)=0 or isnull(dt) then
	     d7date=""
	else
                      xy=right("000"+ cstr((year(dt)-1911)),2)     'è£œé›¶
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      d7date=xy & "/" & xm & "/" & xd
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
	CkStr="'" & s & "'" & endchar
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
	pkStr="'" & s & "'" & endchar
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
	dfn= s & ","
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

function tagChg(s)                        '±N< > ¨ú¥N  2000/12/14
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

function blen(xs)              'ÀË´ú¤¤­^¤å§¨Âø¦r¦ê¹ê»Úªø«×
  xl = len(xs)
  for i=1 to len(xs)
	if asc(mid(xs,i,1))<0 then xl = xl + 1
  next
  blen = xl
end function


Function stripHTML(strHTML)
'Strips the HTML tags from strHTML

 		Dim objRegExp, strOutput
 		Set objRegExp = New Regexp

 		objRegExp.IgnoreCase = True
 		objRegExp.Global = True
 		objRegExp.Pattern = "(<.+?>)|(%3C.+?%3E)|(<.+?%3E)|(%3C.+?>)"

		If IsNull(strHTML) Then strHTML = ""
 		'Replace all HTML tag matches with the empty string
 		strOutput = objRegExp.Replace(strHTML, "")
 
 		'Replace all < and > with &lt; and &gt;       () to ""
 		strOutput = Replace(strOutput, "<", "&lt;")
 		strOutput = Replace(strOutput, ">", "&gt;")
 		strOutput = Replace(strOutput, "%3C", "&lt;")
 		strOutput = Replace(strOutput, "%3E", "&gt;")
		strOutput = Replace(strOutput, "(", "&lt;")
		strOutput = Replace(strOutput, ")", "&gt;")
		strOutput = Replace(strOutput, "%28", "&lt;")
		strOutput = Replace(strOutput, "%29", "&gt;")
 		stripHTML = strOutput 'Return the value of strOutput
 		
 		Set objRegExp = Nothing

End Function

Function pkStrWithSriptHTML(s, endchar)
	pkStrWithSriptHTML=pkStr(stripHTML(s), endchar)
End Function


%>

