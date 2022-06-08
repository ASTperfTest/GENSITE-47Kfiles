<%
'--- 顯示設定字數範圍     s 為目標    l 為字數

function leftstr(s,l)
  if len(s)>l then
    leftstr=left(s,l) & "..."
  else
    leftstr=s
  end if
end function

'--- 抓取 User 換行處 <Memo>

function message(tempstr) 
  outstring = "" 
  while len(tempstr) > 0 
    pos=instr(tempstr, chr(13)&chr(10)) 
    if pos = 0 then 
      outstring = outstring & tempstr & "<br>" 
      tempstr = "" 
    else 
      outstring = outstring & left(tempstr, pos-1) & "<br>" 
      tempstr=mid(tempstr, pos+2) 
    end if 
  wend 
  message = outstring 
end function 

'--- 轉換資料內的單引號

FUNCTION pkStr(s)
  if trim(s)="" then
    pkStr="NULL"
  else
    s=trim(s)
    pos = InStr(s, "'")
    While pos > 0
      s = Mid(s, 1, pos) & "'" & Mid(s, pos + 1)
      pos = InStr(pos + 2, s, "'")
    Wend
    pkStr="'" & s & "'"
  end if
END FUNCTION

'--- 以逗點作分隔之資料轉成陣列

function multi(tstr)
  dim outstr()
  to_no=1
  while len(tstr) > 0
    pos=instr(tstr, ",")
    if pos = 0 then
      redim preserve outstr(to_no)
      outstr(to_no) = tstr
      tstr = ""
    else
      redim preserve outstr(to_no)
      outstr(to_no) = left(tstr, pos-1)
      tstr=trim(mid(tstr, pos+1))
      to_no=to_no+1
    end if
  wend
  multi = outstr
end function

%>