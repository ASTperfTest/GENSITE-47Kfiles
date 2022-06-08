<%
function ServerDate(dt)             'SQL Server 端的系統日期，供Compare時使用
   if Len(dt)=0 Or isnull(dt) then
       ServerDate=""
   else
       ServerDate = Year(now()) & "/" & Month(now()) & "/" & Day(now())
   end if
end function

function dbDate(dt)             'SQL Server 端的系統日期，供Compare時使用
   if Len(dt)=0 Or isnull(dt) then
       dbDate=""
   else
       dbDate = Year(dt) & "/" & Month(dt) & "/" & Day(dt)
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
                      xy=cstr((year(dt)-1911))    '補零
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      d7date=xy & "/" & xm & "/" & xd
                  end if
end function

function s7date(dt)     '轉成民國年  999/99/99 給資料型態為SmallDateTime 使用
	if Len(dt)=0 or isnull(dt) then
	     s7date=""
	else
                      xy=cstr((year(dt)-1911))     '補零
                      xm=right("00"+ cstr(month(dt)),2)
                      xd=right("00"+ cstr(day(dt)),2)
                      s7date=xy & "/" & xm & "/" & xd
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
	dfs="N'" & s & "',"
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
	dfd="N'" & s & "',"
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
	drs="N'" & s & "',"
  end if
END FUNCTION
FUNCTION drs2 (xs)
  s = xUpForm(xs)
  if s="" then
	drs2 = "null,"
  else
	pos = InStr(s, "'")
	While pos > 0
		s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
		pos = InStr(pos + 2, s, "'")
	Wend
	drs2="N'" & s & "',"
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
	drd="N'" & s & "',"
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

FUNCTION drn2 (xs)
s = xUpForm(xs)
  if s="" then
	drn2 = "null,"
  else
	drn2= s & ","
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

function ftpDo(FTPIP,FTPPort,FTPID,FTPPWD,fileAction,FTPfilePath,fileDir,fileTarget,fileSource)	'----FTP機制  2004/7/7

	Set oFtp = CreateObject("FtpCom.FTP")
	oFtp.Connect FTPIP,FTPPort,FTPID,FTPPWD
	
	if left(oFtp.GetMsg,1) = "2" then 
		if FTPfilePath <> "" then oFtp.Execute("CWD " + FTPfilePath)
		if fileAction="CreateDir" then oFtp.CreateDir(fileDir)
		if fileAction="DeleteDir" then oFtp.DeleteDir(fileDir)
		if fileAction="MoveFile" then oFtp.MoveFile fileTarget,fileSource,1,0
		if fileAction="DELE" then oFtp.Execute("DELE " + fileTarget)
		if left(oFtp.GetMsg,1) <> "2" then FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	else
		FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	end if	
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

<script language=VBS>
function blen(xs)              '檢測中英文夾雜字串實際長度
  xl = len(xs)
  for i=1 to len(xs)
	if asc(mid(xs,i,1))<0 then xl = xl + 1
  next
  blen = xl
end function

function xStdTime(dt)  '轉成民國年及時間  999/99/99 10:00 給資料型態為DateTime 使用
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
           d7date=dt
        end if
end function

 function ChkIDX(SID)      '-----身份證字號檢驗

   SID = UCase(SID)
   X1 = InStr("ABCDEFGHJKLMNPQRSTUVWXYZIO", Left(SID, 1)) + 9

   ' 初步檢查 ID 的合法性
   If X1 < 10 Then Chk_ID = False: Exit Function
   If Not IsNumeric(Mid(SID, 2)) Then Chk_ID = False: Exit Function
   If Len(SID) <> 10 Then Chk_ID = False: Exit Function

   ' 檢查編號之正確性
   SID = Cstr(X1) + Mid(SID, 2)

   nCheckSum = Cint(Mid(SID, 1, 1))
   For I = 2 To Len(SID) - 1
        nCheckSum = nCheckSum + Cint(Mid(SID, I, 1)) * (11 - I)
   Next
   nCheckSum = nCheckSum + Cint(Mid(SID, 11, 1))

   ChkID = (nCheckSum Mod 10 = 0)
end Function

function ChkID(idno)	'-----身份證字號檢驗(新)
  if len(trim(idno))=0 then exit function

  alpha=UCase(left(idno,1))
  d1=mid(idno,2,1)
  d2=mid(idno,3,1)
  d3=mid(idno,4,1)
  d4=mid(idno,5,1)
  d5=mid(idno,6,1)
  d6=mid(idno,7,1)
  d7=mid(idno,8,1)
  d8=mid(idno,9,1)
  d9=mid(idno,10,1)
  select case alpha
    case "A" : acc=1
    case "B" : acc=10
    case "C" : acc=19
    case "D" : acc=28
    case "E" : acc=37
    case "F" : acc=46
    case "G" : acc=55
    case "H" : acc=64
    case "I" : acc=39
    case "J" : acc=73
    case "K" : acc=82
    case "L" : acc=2
    case "M" : acc=11
    case "N" : acc=20
    case "O" : acc=48
    case "P" : acc=29
    case "Q" : acc=38
    case "R" : acc=47
    case "S" : acc=56
    case "T" : acc=65
    case "U" : acc=74
    case "V" : acc=83
    case "W" : acc=21
    case "X" : acc=3
    case "Y" : acc=12
    case "Z" : acc=30
  end select
  on error resume next
  checksum = acc+8*d1+7*d2+6*d3+5*d4+4*d5+3*d6+2*d7+1*d8+1*d9
  check1 = Int(checksum/10)
  check2 = checksum/10
  check3 = (check2-check1)*10
  if len(idno)>10 then
    ChkID=false
  elseif err.number<>0 then
    ChkID=false
  elseif checksum=check1*10 then
    ChkID=true
  elseif d9=(10-check3) then
    ChkID=true
  else
    ChkID=false
  end if
end function
function ftpDo(FTPIP,FTPPort,FTPID,FTPPWD,fileAction,FTPfilePath,fileDir,fileTarget,fileSource)	'----FTP機制  2004/7/7
response.end
	Set oFtp = CreateObject("FtpCom.FTP")
	oFtp.Connect FTPIP,FTPPort,FTPID,FTPPWD
	
	if left(oFtp.GetMsg,1) = "2" then 
		if FTPfilePath <> "" then oFtp.Execute("CWD " + FTPfilePath)
		if fileAction="CreateDir" then oFtp.CreateDir(fileDir)
		if fileAction="DeleteDir" then oFtp.DeleteDir(fileDir)
		if fileAction="MoveFile" then oFtp.MoveFile fileTarget,fileSource,1,0
		if fileAction="DELE" then oFtp.Execute("DELE " + fileTarget)
		if left(oFtp.GetMsg,1) <> "2" then FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	else
		FTPErrorMSG="  FTP機制出現錯誤,FTP未成功!"
		oFtp.LogOffServer 
		set oFtp = nothing
	end if	
end function

</script>