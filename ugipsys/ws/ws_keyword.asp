<?xml version="1.0" encoding="utf-8"?>
<xKeywordList>
<%
function blen(xs)              '檢測中英文夾雜字串實際長度
  xl = len(xs)
  for p=1 to len(xs)
	if asc(mid(xs,p,1))<=0 then xl = xl + 1
  next
  blen = xl
end function

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
	xKeyword=pkStr(request.querystring("xKeyword"),"")
	xKeyword=mid(xKeyword,2)
	xKeyword=left(xKeyword,len(xKeyword)-1)
	xStr=""
	xReturnValue=""
	xKeywordArray=split(xkeyword,",")
	
	for i=0 to ubound(xKeywordArray)
		'----去除權重符號
		xPos=instr(xKeywordArray(i),"*")
		if xPos<>0 then
			xStr=Left(trim(xKeywordArray(i)),xPos-1)
		else
			xStr=trim(xKeywordArray(i))
		end if
		
		'----叫用hysearch, 若xStr不存在串入字串中
		' step 1: 使用 CreateObject 產生 Object. DLL name: hysdk, class name: hyft  
		Set obj = Server.CreateObject("hysdk.hyft.1")

		' step 2: 呼叫 hyft_connect 連結 hyftd	hyft_connect() 傳回值為連線 id, 存放於 obj.ret_val 中
		idxStr = session("mySiteID")
		x = obj.hyft_connect("localhost", 2816, idxStr , "", "")
		id = obj.ret_val

		query=xStr

		' step 3: 啟用hyft_textkeys斷詞功能,如果成功,會得到一個key_id,
		x=obj.hyft_textkeys(id, query,10)
		key_id = obj.ret_val

		' step 4: 算出query中斷詞的數目
		x=obj.hyft_num_keys(id,key_id)
		num =obj.ret_val
		
		x=obj.hyft_keylen(id,key_id,0)
		key_len=obj.ret_val
	
		x=obj.hyft_free_key(id,key_id)

		set obj=Nothing			
		'----hysearch完成
	   	response.write "<Hello>"+request.querystring("xKeyword")+"</Hello>"
		response.write "<Hello5>"+query+"</Hello5>"		
		response.write "<Hello2>"+cStr(num)+"</Hello2>"
		response.write "<Hello3>"+cStr(key_len)+"</Hello3>"
		response.write "<Hello4>"+cStr(blen(xStr))+"</Hello4>"			
		if num=0 or key_len<>blen(xStr) then xReturnValue=xReturnValue+xStr+";"
	next
	if xReturnValue<>"" then xReturnValue=Left(xReturnValue,Len(xReturnValue)-1)
	response.write "<xKeywordStr>"+xReturnValue+"</xKeywordStr>"
%>
</xKeywordList>

