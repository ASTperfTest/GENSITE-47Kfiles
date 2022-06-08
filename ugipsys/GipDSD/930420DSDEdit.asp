<%@ CodePage = 65001 %>
<%
Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'Conn.Open session("ODBCDSN")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = session("ODBCDSN")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	fSql = "sp_tables"
	set RSTable = conn.execute(fSql)
	TableArray=RSTable.getrows()
	for i=0 to ubound(TableArray,2)
	    if (Left(TableArray(2,i),5)="CuDTx" and TableArray(2,i)<>"CuDTxMMO") or (Left(TableArray(2,i),2)="cm" and len(TableArray(2,i))=4) then
	    	SQLUpdate=""
		fSql = "sp_columns '" & TableArray(2,i) & "'"
		set RSlist = conn.execute(fSql)
		while not RSlist.eof
		   if RSlist("COLUMN_NAME")<>"giCuItem" and RSlist("NULLABLE")=0 then
		   	if RSlist("TYPE_NAME")="char" or RSlist("TYPE_NAME")="varchar" then
				SQLUpdate=SQLUpdate+"Alter Table "+TableArray(2,i)+" Alter Column "+RSlist("COLUMN_NAME")+" "+RSlist("TYPE_NAME")+"("+cStr(RSlist("LENGTH"))+")"+ " NULL;"	    
'				SQLUpdate="Alter Table "+TableArray(2,i)+" Alter Column "+RSlist("COLUMN_NAME")+" "+RSlist("TYPE_NAME")+"("+cStr(RSlist("LENGTH"))+")"+ " NULL;"	    
'				response.write SQLUpdate+"<br>"
		   	elseif RSlist("TYPE_NAME")="int" or RSlist("TYPE_NAME")="datetime" or RSlist("TYPE_NAME")="smalldatetime" then
				SQLUpdate=SQLUpdate+"Alter Table "+TableArray(2,i)+" Alter Column "+RSlist("COLUMN_NAME")+" "+RSlist("TYPE_NAME")+ " NULL;"	    
'				SQLUpdate="Alter Table "+TableArray(2,i)+" Alter Column "+RSlist("COLUMN_NAME")+" "+RSlist("TYPE_NAME")+ " NULL;"	    
'				response.write SQLUpdate+"<br>"
			end if
		    end if
		    RSlist.movenext
		wend
		if SQLUpdate<>"" then conn.execute(SQLUpdate)
'		response.write SQLUpdate+"<hr>"
	    end if	
	next

response.write "DONE2!"
response.end

%>
