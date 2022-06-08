<?xml version="1.0"  encoding="utf-8" ?>
<divList>
<%
 Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
' conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


 catCode = left(request("locCode"),1)    ' -- 防 SQL injection 用
 sql = "SELECT mCode, mValue FROM codeMain WHERE codeMetaID=N'locArea" & catCode & "'"
 'CuDTx39F155 表銀行所在區域
 'sql = "SELECT mCode, mValue FROM codeMain WHERE codeMetaID=N'CuDTx39F155" & catCode & "'"
 set RS = conn.execute(sql)
 while not RS.eof
  response.write "<row><mCode>" & RS("mCode") & "</mCode><mValue>" &
RS("mValue") & "</mValue></row>" & vbCRLF
  RS.moveNext
 wend

%>
</divList>
