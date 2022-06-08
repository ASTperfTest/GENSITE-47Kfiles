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

function d7date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
	     d7date=""
	else
                      xy=right("000"+ cstr((year(dt)-1911)),3)     '補零
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

function CheckAndReplaceID( id ) 
	id = Trim(id)
    dim tab
	A1=array(1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,3,3,3,3,3,3 )
	A2=array(0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5 )
	Mx=array(9,8,7,6,5,4,3,2,1,1)
	tab = "ABCDEFGHJKLMNPQRSTUVXYWZIO" 		  
	if  Len(id) <> 10 then 
		CheckAndReplaceID = id
		exit function
	end if
	i = InStr(tab,Left(id,1)) 
	if  i = 0 then
		CheckAndReplaceID = id
		exit function
	end if 
	sum = A1(i) + A2(i)*9

	for  i=2 to 10 
		temp = Mid(id,i,1)
		if not IsNumeric(temp) then
			CheckAndReplaceID = id
			exit function
		end if
		v = CInt(temp)
		sum = sum +  v*Mx(i-1)
	next
	if sun mod 10 <> 0 then 
		CheckAndReplaceID = id
		exit function
	end if
	
	CheckAndReplaceID = Left(id,1) + "*****" + Right(id,4)
end function

function ReplaceEmailId(email)
	dim temp 
	i = InStr(email,"@") 
	if i <> 11 then
		ReplaceEmailId = email
		exit function
	end if
	temp = Left(email,10)
	temp = CheckAndReplaceID(temp)
	
	ReplaceEmailId = temp + Mid(email,10,Len(email))
end function

%>

