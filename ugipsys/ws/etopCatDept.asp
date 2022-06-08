<?xml version="1.0"  encoding="utf-8" ?>
<divList>
<%
	session("ODBCDSN")="Provider=SQLOLEDB;Data Source=10.10.5.128;User ID=hyGIP;Password=hyweb;Initial Catalog=GIPmof"
'	session("ODBCDSN")="Provider=SQLOLEDB;Data Source=210.69.109.16;User ID=hyMOF;Password=mof0530;Initial Catalog=mofgip"
	Set conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open session("ODBCDSN")
Set conn = Server.CreateObject("HyWebDB3.dbExecute")
conn.ConnectionString = session("ODBCDSN")
'----------HyWeb GIP DB CONNECTION PATCH----------


	catCode = request("catCode")
	sql = "SELECT deptID, eAbbrName, xURL FROM dept AS d JOIN CuDTGeneric AS g" _
		& " ON d.giCuItem=g.iCuItem WHERE xURL<>'' AND tDataCat LIKE N'%" & catCode & "%'"
	set RS = conn.execute(sql)
	while not RS.eof
		response.write "<row><mCode>" & RS("xURL") & "</mCode><mValue>" & RS("eAbbrName") & "</mValue></row>" & vbCRLF
		RS.moveNext
	wend
	
%>
</divList>
