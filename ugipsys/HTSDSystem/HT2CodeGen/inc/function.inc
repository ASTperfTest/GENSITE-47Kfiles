<%
'--- ��ܳ]�w�r�ƽd��     s ���ؼ�    l ���r��

function leftstr(s,l)
  if len(s)>l then
    leftstr=left(s,l) & "..."
  else
    leftstr=s
  end if
end function

'--- ��� User ����B <Memo>

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

'--- �ഫ��Ƥ�����޸�

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

'--- �H�r�I�@���j������ন�}�C

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