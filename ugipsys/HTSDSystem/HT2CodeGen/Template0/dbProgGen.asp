<%@ Language=VBScript %>
<!-- #INCLUDE FILE="adovbs.inc" -->
<%

if request("task")="SET ODBC" then
	session("odbcStr") ="DSN="&request("odbcStr")&";UID=hometown;PWD=2986648;"
end if 

%>
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<FORM name=ooo method=post>
ODBC：
<INPUT name=odbcStr size=20 value="<%=request("odbcStr")%>"> 
<INPUT type=submit name=task value="SET ODBC">　　

<% if session("odbcStr") <> "" then %>
<SELECT name=fxTblName size=1>
<%
    Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open Session("odbcStr")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = Session("odbcStr")
'----------HyWeb GIP DB CONNECTION PATCH----------

	

'----------HyWeb GIP DB CONNECTION PATCH----------
'	set RS = conn.OpenSchema (adSchemaTables)
Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = Schema (adSchemaTables)
'----------HyWeb GIP DB CONNECTION PATCH----------
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
	while not RS.eof
		if RS("TABLE_TYPE") = "TABLE" then
%>
<OPTION value="<%=RS("TABLE_NAME")%>" <% if request("fxTblName")=RS("TABLE_NAME") then Response.Write " selected"%>><%=RS("TABLE_NAME")%></OPTION>
<%
		end if
		RS.moveNext
	wend
%>
</SELECT>
<INPUT type=submit name=task value="SELECT TABLE">
<HR>
</FORM>
<%
  if request("task")="SELECT TABLE" AND request("fxTblName")<>"" then
%>
<CENTER>
<FORM name=f1 METHOD="POST" ACTION="dbProgGo.asp">
<TABLE BORDER><THEAD>
<TR><TD>欄位名稱<TD>描述<td>型別<td>長度<td width=20%>編輯碼<td>查詢條列
<TBODY>
<%  

'----------HyWeb GIP DB CONNECTION PATCH----------
'	set RS = conn.OpenSchema (adSchemaColumns)
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = Schema (adSchemaColumns)
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	
	while not RS.eof
		if RS("TABLE_NAME") = request("fxTblName") then
			select case RS("DATA_TYPE")
			  case 2:	dfType = "SmallInt"
			  case 3:	dfType = "Integer"
			  case 4:	dfType = "Single"
			  case 5:	dfType = "Double"
			  case 6:	dfType = "Currency"
			  case 7:	dfType = "Date"
			  case 8:	dfType = "BSTR"
			  case 9:	dfType = "IDispatch"
			  case 10:	dfType = "Error"
			  case 11:	dfType = "Boolean"
			  case 12:	dfType = "Variant"
			  case 13:	dfType = "IUnknown"
			  case 14:	dfType = "Decimal"
			  case 16:	dfType = "TinyInt"
			  case 17:	dfType = "UnsignedTinyInt"
			  case 18:	dfType = "UnsignedSmallInt"
			  case 19:	dfType = "UnsignedInt"
			  case 20:	dfType = "BigInt"
			  case 21:	dfType = "UnsignedBigInt"
			  case 64:	dfType = "FileTime"
			  case 72:	dfType = "GUID"
			  case 128:	dfType = "Binary"
			  case 129:	dfType = "Char"
			  case 130:	dfType = "WChar"
			  case 131:	dfType = "Numeric"
			  case 132:	dfType = "UserDefined"
			  case 133:	dfType = "DBDate"
			  case 134:	dfType = "DBTime"
			  case 135:	dfType = "DBTimeStamp"
			  case 136:	dfType = "Chapter"
			  case 137:	dfType = "DBFileTime"
			  case 138:	dfType = "PropVariant"
			  case 139:	dfType = "VarNumeric"
			  case 200:	dfType = "VarChar"
			  case 201:	dfType = "LongVarChar"
			  case 202:	dfType = "VarWChar"
			  case 203:	dfType = "LongVarWChar"
			  case 204:	dfType = "VarBinary"
			  case 205:	dfType = "LongVarBinary"
			  case else: dfType = "Unknown"
			end select
%>
<TR>
  <TD><%=RS("COLUMN_NAME")%>   
  <%if  isnull(RS("DESCRIPTION")) then %>
         <TD><%=RS("COLUMN_NAME")%>
  <%else%>
         <TD><%=RS("DESCRIPTION")%>
  <%end if%>  
  
  <TD><%=RS("DATA_TYPE")%>-<%=dfType%>
  <TD>
<%
x=RS("CHARACTER_MAXIMUM_LENGTH")
if x > 255 then x = "MEMO"
response.write x

	preCode = "t"	
	if RS("DATA_TYPE")=135 then 
		preCode = "d"
	elseif RS("DATA_TYPE")=11 then
		preCode = "b"
	elseif RS("DATA_TYPE")<7 OR (RS("DATA_TYPE")>13 AND RS("DATA_TYPE")<22) OR RS("DATA_TYPE")=131 then
		preCode = "n"
	end if
%>
<TH><SELECT name=prefix_<%=RS("COLUMN_NAME")%> size=1>
<OPTION value="t" <%if preCode="t" then %>selected<%end if%>>T-文字</OPTION>
<OPTION value="p" <%if preCode="p" then %>selected<%end if%>>P-主鍵</OPTION>
<OPTION value="n" <%if preCode="n" then %>selected<%end if%>>N-數值</OPTION>
<OPTION value="d" <%if preCode="d" then %>selected<%end if%>>D-日期</OPTION>
<OPTION value="s" <%if preCode="s" then %>selected<%end if%>>S-選取</OPTION>
<OPTION value="b" <%if preCode="b" then %>selected<%end if%>>B-布林</OPTION>
<OPTION value="" <%if preCode="" then %>selected<%end if%>>--空--</OPTION>
</SELECT>
<SPAN id=CodeID<%=RS("COLUMN_NAME")%> style="display:none"><Select name=CodeIDSelect<%=RS("COLUMN_NAME")%> size=1>
</select></SPAN>
<!--<INPUT name=prefix_<%=RS("COLUMN_NAME")%> size=1 value=<%=preCode%> style="text-align:center; font-size:100%;">-->
<TH><INPUT type=checkbox name=doList_<%=RS("COLUMN_NAME")%>>

<%
		end if
		RS.moveNext
	wend
%>
</TABLE>
<INPUT type=hidden name=tableName value="<%=request("fxTblName")%>">
<TABLE>
<TR><TD>程式權限代碼：<TD><INPUT name=progRightCode size=10 value="">
<TR><TD>程式前置名稱：<TD><INPUT name=programPrefix size=16 value="<%=request("fxTblName")%>">
<TR><TD>中文資料名稱：<TD><INPUT name=progCapPrefix size=16 value="ＸＸ資料">
<TR><TD>程式置放路徑：<TD><INPUT name=programPath size=30 value="伺服端的絕對路徑">
<tr><th colspan=2><INPUT type=submit value="產生程式">　　<INPUT type=reset>
</TABLE>
</FORM>
</BODY>
</HTML>
<script language=VBS>
redim codeIDArray(2,0)
<%
SQL="Select CodeID,CodeName,CodeTblName from CodeTable Order By CodeTblName,CodeID"
SET RSCode=conn.execute(SQL)
while not RSCode.EOF
    i=i+1%>
    redim preserve codeIDArray(2,<%=i%>)
    codeIDArray(0,<%=i%>)="<%=RSCode("CodeID")%>"
    codeIDArray(1,<%=i%>)="<%=RSCode("CodeName")%>"
    codeIDArray(2,<%=i%>)="<%=RSCode("CodeTblName")%>"        
<%  RSCode.movenext
wend
%>
for i=1 to <%=i%>
'alert codeIDArray(0,i) & "[]" & codeIDArray(1,i) & "[]" & codeIDArray(2,i)
next
sub document_onClick
    set sObj=window.event.srcElement
    if sObj.tagname="SELECT" then
       if left(sObj.name,7)="prefix_" then
	  xcName=mid(sObj.name,8)
	  if sObj.value="s" then  
	     set xsrc = document.all("CodeIDSelect"&xcName)
		 xaddOption xsrc,"請選擇",""
	     for k=1 to <%=i%>
		 xaddOption xsrc,codeIDArray(0,k),codeIDArray(0,k)
	     next
     	     xsrc.selectedIndex = 0
	     document.all("CodeID"&xcName).style.display="block"     	     	     	     
	  else
	     document.all("CodeID"&xcName).style.display="none"
	  end if   
       end if	
    end if	
end sub

sub xaddOption(xlist,name,value)  
	set xOption = document.createElement("OPTION")         
	xOption.text=name         
	xOption.value=value              
	xlist.add xOption
end sub 	
</script>
<%  
  end if
end if
%>
