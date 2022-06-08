<!--#INCLUDE FILE="../../inc/rtf.inc" -->
<% HTProgCap="Code代碼定義表"
HTProgCode="代碼定義表"
HTProgPrefix="Code" %>
<%
Response.Expires = 0

dim rs, sql
dim rtf, rtf_head, rtf_tail

	Set conn = Server.CreateObject ("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
'Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
conn.ConnectionTimeout=0
conn.CursorLocation = 3
conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	set rs=server.CreateObject("ADODB.Recordset")
	
        SQL="Select * from CodeTable where CodeID='" & request.querystring("codeID") & "'"
        set RScode = conn.execute(sql)
        
        fSql=""  
        fsql2=""

        if isNull(RSCode("CodeSrcItem")) then
	    if isNull(RSCode("CodeSortFld")) then    
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & " From " & RSCode("CodeTblName")
	    else
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & " From " & RSCode("CodeTblName") & " Order By " & RSCode("CodeValueFld")	
	    end if		
        else
	    if isNull(RSCode("CodeSortFld")) then
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & " From " & RSCode("CodeTblName") & " where " & RSCode("CodeSrcFld") & "='" & RSCode("CodeSrcItem") & "'"	
	    else    
		fsql="Select " & RSCode("CodeValueFld") & "," & RSCode("CodeDisplayFld") & "," & RSCode("CodeSortFld") & " From " & RSCode("CodeTblName") & " where " & RSCode("CodeSrcFld") & "='" & RSCode("CodeSrcItem") & "' Order By " & RSCode("CodeSortFld")
	    end if
        end if
        
        if request("xOption")="2" then
                fsql2=RSCode("CodeValueFld") &" between '" &request("ValueFm") &"' and '" &request("ValueTo") &"'" 
        end if
        if request("xOption")="3" then
                fsql2=RSCode("CodeSortFld") &" between '" &request("SortFm") &"' and '" &request("SortTo") &"'" 
        end if
        
    	if fsql2<>"" then 
		pos1=instr(fSql,"where")
        	pos2=instr(fSql,"Order By")
		if pos1=0 and pos2=0 then
	    		fSql=fSql & " where " & fsql2
	    	elseif pos1=0 and pos2<>0 then
		        fSql=left(fSql,pos2-1) & "where " & fsql2 & mid(fSql,pos2-1)
    		elseif pos1<>0 and pos2=0 then
	    		fSql=fSql & " AND " & fsql2 
    		else
	        	fSql=left(fSql,pos2-1) & " AND " & fsql2 & mid(fSql,pos2-1)
    		end if   	
    	end if
   
       if instr(fSql,"Order By")=0 then
		fsql = fsql & " Order By " & RSCode("CodeValueFld")  
       end if
       
'response.write fsql       
       set rs=conn.execute(fsql)
        

	if not rs.EOF then
		Response.ContentType = "application/msword"	
		
		rtf = rtf_open(HTProgPrefix & ".rtf")
		par(0) = rtf_readpar(HTProgPrefix & "Par0.rtf")
		
		find_head_tail rtf, rtf_head, rtf_tail
		
		head=rtf_head
		rtf_field head,1,"\f27"&"\fs35"&"代碼:"&request("CodeID")&request("CodeName")
		rtf_field head,2,"\f27"&RScode("CodeValueFldName")
		rtf_field head,3,"\f27"&RScode("CodeSortFldName")
		Response.Write head
		
		i=0   
                while not RS.EOF
                     i=i+1
                     if ISNULL(RScode("CodeSortFld")) then
                            response.write rtf_addrow (0,i &";" &rs(0) &";" &"\f27"&rs(1) &";" )
                     else 
                            response.write rtf_addrow (0,i &";" &rs(0) &";" &"\f27"&rs(1) &";" &rs(2))
                     end if
                        
                     Rs.movenext
                wend  
                
		Response.Write rtf_tail  
		        
		       
		
	else %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=HTProgCap%></title>
</head>
<body>
<script language=vbs>
	on error resume next
	alert("找不到資料")
	window.history.back
</script>
</body>
</html>
<%  end if %>
<%
function YearName(year)
	dim y, d1, d2, d3, c, ret
	
	if isnull(year) then
		YearName = ""
		exit function
	end if
	
	y = cint(year)
	d1 = y \ 100
	d2 = (y - d1 * 100) \ 10
	d3 = (y - d1 * 100 - d2 * 10)
		
	c = "零一二三四五六七八九"
	ret = ""
	if d1 > 0 then
		ret = mid(c, d1+1, 1) & "百"
	end if
	if d2 > 0 then
		ret = ret & mid(c, d2+1, 1) & "十"
	elseif d3 > 0 then
		ret = ret & mid(c, d2+1, 1)
	end if
	if d3 > 0 then
		ret = ret & mid(c, d3+1, 1)
	end if
	
	YearName = ret	
end function

function ChMoney(n)
        if isnull(n) then
		ChMoney = ""
		exit function
	end if
	c = "○一二三四五六七八九"
	ret = ""
	x=len(n)
	i=0
	while not x=0
	      i=i+1
	      ret=ret & mid(c,mid(n,i,1)+1,1)
	      x=x-1
	wend
	chMoney=ret 
end function
%>
