<%@ CodePage = 65001 %>
<% 
Response.Expires = 0 
Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

Server.ScriptTimeOut = 2000

'----將資料表欄位型態放入陣列中
redim colname(2,0)
i=0
tableName = request("tbl")
fSql = "sp_columns " & pkStr(tableName,"")
set RSList = conn.execute(fSql)
if not RSList.EOF then
    while not RSList.EOF
    	redim preserve colname(2,i)
    	colname(0,i)=i: colname(1,i)=RSList("TYPE_NAME"): colname(2,i)=RSList("COLUMN_NAME")
    	i=i+1
    	RSList.movenext
    wend 
end if
'for j=0 to ubound(colname,2)
'    response.write cStr(colname(0,j))+"[]"+colname(1,j)+"[]"+colname(2,j)+"<br>"
'next
'response.end
'----取出資料表資料並串成字串
cvbCRLF = "<BR>" & vbCRLF
SQL="Select * from " & tableName
SET RSD=conn.execute(SQL)
SQLIns=""
if not RSD.EOF then
    while not RSD.EOF
        SQLInsStr1="Insert Into "&tableName&" ("
        SQLInsStr2=" Values("        
	for k=0 to ubound(colname,2)
	    if not ISNULL(RSD(colname(0,k))) then
	    	if instr(colname(1,k),"identity")>0 then
	    	    SQLInsStr1 = SQLInsStr1 + "xi_" + colname(2,k) + ","
	    	    SQLInsStr2 = SQLInsStr2 + cStr(RSD(colname(0,k))) + ","
	    	else
	    	    SQLInsStr1 = SQLInsStr1 + colname(2,k) + ","
	    	    SQLInsStr2 = SQLInsStr2 + StrCombine(RSD(colname(0,k)),colname(1,k))
	    	end if
	    end if
	next
	SQLInsStr=left(SQLInsStr1,len(SQLInsStr1)-1) + ")" + left(SQLInsStr2,len(SQLInsStr2)-1) + ");" + cvbCRLF	
	response.write SQLInsStr
    	RSD.movenext
    wend 
end if

response.end

'----字串組合檢查欄位型態函數	
function StrCombine(colvalue,coltype)
	xStr=""
	if coltype="int" or coltype="tinyint" or coltype="smallint" or coltype="smallint" or coltype="float" then
		xStr=cStr(colvalue)+","
	else
		xStr=pkStr((colvalue),",")	
	end if
	StrCombine=xStr
end function

'----quote處理
FUNCTION pkStr (s, endchar)
  if ISNULL(s) then
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
%>
