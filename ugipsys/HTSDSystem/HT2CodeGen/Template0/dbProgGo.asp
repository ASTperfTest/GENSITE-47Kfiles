<%@ Language=VBScript %>
<!-- #INCLUDE FILE="adovbs.inc" -->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
     td.fldLable {text-align:right; font-size:smaller;}
</STYLE>
</HEAD>
<BODY>
<%
Function SelectOptionList(codeID)
	SQL="Select * from CodeTable where codeID='" & codeID & "'"
        SET RSLK=conn.execute(SQL)
	str=""
        if not RSLK.EOF then
          if isNull(RSLK("CodeSortFld")) then
           if isNull(RSLK("CodeSrcFld")) then
	    	str=ct&ct&ct&"<option value="""" style=""color:blue"">請選擇</option>"&VBCRLF& _
		    ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & """" & VBCRLF & _
		    ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
		    ct&ct&ct&"if not RSS.EOF then"& VBCRLF & _
		    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF & _
		    ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"& VBCRLF & _
		    ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
		    ct&ct&ct&"wend"& VBCRLF & _
		    ct&ct&ct&"end if"&cr	               
           else	
	    	str=ct&ct&ct&"<option value="""" style=""color:blue"">請選擇</option>"&VBCRLF& _
		    ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "'""" & VBCRLF & _
		    ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
		    ct&ct&ct&"if not RSS.EOF then"& VBCRLF & _
		    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF & _
		    ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"& VBCRLF & _
		    ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
		    ct&ct&ct&"wend"& VBCRLF & _
		    ct&ct&ct&"end if"&cr	    
	   end if          
          else
           if isNull(RSLK("CodeSrcFld")) then
	    	str=ct&ct&ct&"<option value="""" style=""color:blue"">請選擇</option>"&VBCRLF& _
		    ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
		    ct&ct&ct&"if not RSS.EOF then"& VBCRLF & _
		    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF & _
		    ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"& VBCRLF & _
		    ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
		    ct&ct&ct&"wend"& VBCRLF & _
		    ct&ct&ct&"end if"&cr	               
           else	
	    	str=ct&ct&ct&"<option value="""" style=""color:blue"">請選擇</option>"&VBCRLF& _
		    ct&ct&ct&cl & "SQL=""Select " & RSLK("CodeValueFld") & "," & RSLK("CodeDisplayFld") & " from " & RSLK("CodeTblName") & " where " & RSLK("CodeSortFld") & " IS NOT NULL AND " & RSLK("CodeSrcFld") & "='" & RSLK("CodeSrcItem") & "' Order by " & RSLK("CodeSortFld") & """" & VBCRLF & _
		    ct&ct&ct&"SET RSS=conn.execute(SQL)" & VBCRLF & _
		    ct&ct&ct&"if not RSS.EOF then"& VBCRLF & _
		    ct&ct&ct&"While not RSS.EOF"&cr& VBCRLF & _
		    ct&ct&ct&"<option value="""&cl&"=RSS(0)"&cr&""">"&cl&"=RSS(1)"&cr&"</option>"& VBCRLF & _
		    ct&ct&ct&cl&ct&"RSS.movenext"& VBCRLF & _
		    ct&ct&ct&"wend"& VBCRLF & _
		    ct&ct&ct&"end if"&cr	    
	   end if
	  end if
	  SelectOptionList=str
        end if
End Function

for each x in request.form
'	response.write x & "==>" & request(x) & "<BR>"
next

pgPrefix = request("programPrefix")
pgPath = request("programpath")
If right(pgPath,1) <> "\" then pgPath = pgPath & "\"
   %>
   <TABLE width=100%>  
   <TR><TD>ODBC：<%=Session("odbcStr")%>　　<TD>TableName：<%=request("tableName")%>
   <TR><TD colspan=2>Program： <%=pgPath%><%=pgprefix%>Query.asp　　<%=request("progCapPrefix")%>管理
   </table>
   <HR>
   <%
    Set Conn = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'	conn.Open Session("odbcStr")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = Session("odbcStr")
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------


	Dim xSearchListItem(20,2)
	xItemCount = 0

    Set fs = CreateObject("scripting.filesystemobject")
'--xxxFORM.inc--------------------------
    Set xfin = fs.opentextfile(Server.MapPath("TempForm1.inc"))
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Form.inc")

	dumpTempFile
    xfin.Close


'----------HyWeb GIP DB CONNECTION PATCH----------
'	set RS = conn.OpenSchema (adSchemaColumns)
Set Conn = Server.CreateObject("ADODB.Connection")
'Set Conn = Server.CreateObject("HyWebDB3.dbExecute")
Conn.ConnectionString = Schema (adSchemaColumns)
Conn.ConnectionTimeout=0
Conn.CursorLocation = 3
Conn.open
'----------HyWeb GIP DB CONNECTION PATCH----------

	regInitStr = ""	
	regOrgStr = ""
	cq=chr(34)
	ct=chr(9)
	cl="<" & "%"
	cr="%" & ">"
	response.write "<TABLE border>"
	xfout.writeline "<CENTER><TABLE border=0 class=bluetable cellspacing=1 cellpadding=2 width=""80%"">"

	While not RS.eof
		if RS("TABLE_NAME") = request("tableName") then
		    if isnull(RS("DESCRIPTION")) then 
		       fDescript=RS("COLUMN_NAME")
		    else    
			   fDescript=RS("DESCRIPTION")
		    end if
		    
			myFldName = RS("COLUMN_NAME")
		   	xpre = LCase(trim(request("prefix_" & RS("COLUMN_NAME"))))
			if request("doList_"&myFldName) = "on" then
				xItemCount = xitemCount + 1
		     	    	xSearchListItem(xItemCount,0) = fDescript
				xSearchListItem(xItemCount,1) = myFldName
				xSearchListItem(xItemCount,2) = xpre
			end if
		'----------------------------------------------------------------------------	
            if xpre<>"" then
		    	xname = xpre & "fx_" & RS("COLUMN_NAME")
		    	xlen=RS("CHARACTER_MAXIMUM_LENGTH")
	    		if RS("DATA_TYPE")=135 then 
		    		xlen=10
		    	elseif RS("DATA_TYPE")<7 OR (RS("DATA_TYPE")>13 AND RS("DATA_TYPE")<22) OR RS("DATA_TYPE")=131 then
			    	xlen=8
		    	end if
		    	if RS("DATA_TYPE")=11 then
			    	regInitStr = regInitStr & ct & "reg." & xname & ".checked = " & cq & cl & "=qqRS(""" & RS("COLUMN_NAME") & """)" & cr & cq & vbCRLF
			    	regOrgStr = regOrgStr & ct & "reg." & xname & ".checked = " & cq & cl & "=request(""reg." & xname & """)" & cr & cq & vbCRLF
			    else
		    		regInitStr = regInitStr & ct & "reg." & xname & ".value = " & cq & cl & "=qqRS(""" & RS("COLUMN_NAME") & """)" & cr & cq & vbCRLF
			    	regOrgStr = regOrgStr & ct & "reg." & xname & ".value = " & cq & cl & "=request(""reg." & xname & """)" & cr & cq & vbCRLF
			    end if
	            xfout.writeline ct & ct & "<TR>"
             	xfout.writeline ct & ct & "  <TD class=lightbluetable align=right>" & fDescript & "：</TD>"			
                %>
		      <TR>
		         <TD class=lightbluetable align=right><%=fDescript%>：</TD><% if xlen > 255 then 
	                    xfout.writeline ct & ct & "  <TD class=whitetablebg><TEXTAREA NAME=" & xname & " ROWS=7 COLS=60></TEXTAREA></TD>" %>
		         <TD class=whitetablebg><TEXTAREA NAME=<%=xname%> ROWS=7 COLS=60></TEXTAREA></TD>
       <%elseif RS("DATA_TYPE")=11 then 
	               xfout.writeline ct & ct & "  <TD class=whitetablebg><INPUT TYPE=checkbox NAME=" & xname & " value=""1""></TD>" %>
		          <TD class=whitetablebg><INPUT TYPE=checkbox NAME=<%=xname%> value="1"></TD>
       <%elseif xpre = "s" then
	              xfout.writeline ct & ct & "  <TD class=whitetablebg><SELECT NAME=" & xname & " SIZE=1>"
	              xfout.writeline SelectOptionList(request("CodeIDSelect"&RS("COLUMN_NAMe")))
	              xfout.writeline ct & ct & ct & "</SELECT></TD>" %>
		         <TD class=whitetablebg><SELECT NAME=<%=xname%> SIZE=1>
			     </SELECT></TD>
       <%elseif xpre = "p" then
	              xfout.writeline ct & ct & "  <TD class=whitetablebg><INPUT TYPE=text NAME=" & xname & " SIZE=" & xlen & cl & "IF request.querystring(""" & myfldName & """) <> """" then " & cr & " class=sedit readonly" & cl & " End If " & cr & "></TD>" %>
		         <TD class=whitetablebg><INPUT TYPE=text NAME=<%=xname%> SIZE=<%=xlen%>></TD>
       <%elseif xpre = "n" then
                 xfout.writeline ct & ct & "  <TD class=whitetablebg><INPUT TYPE=text NAME=" & xname & " SIZE=" & xlen & " style=""text-align:right;""></TD>" %>
		        <TD class=whitetablebg><INPUT TYPE=text NAME=<%=xname%> SIZE=<%=xlen%> style="text-align:right;"></TD>
       <%else
	            xfout.writeline ct & ct & "  <TD class=whitetablebg><INPUT TYPE=text NAME=" & xname & " SIZE=" & xlen & "></TD>" %>
		        <TD class=whitetablebg><INPUT TYPE=text NAME=<%=xname%> SIZE=<%=xlen%>></TD>
       <%end if %></TR>
     '---------------------------------------------------------------------------------------------  
<%
            	xfout.writeline ct & ct & "</TR>"
		  end if
		  
		end if
		RS.moveNext
	Wend
	
	response.write "</TABLE>"
	xfout.writeline "</TABLE></CENTER>"

    Set xfin = fs.opentextfile(Server.MapPath("TempForm2.inc"))
	dumpTempFile
    xfin.Close
    xfout.Close

%>
<!-- 讀取資料庫，設定初值常式：
<%=regInitStr%>
-->
<!-- 讀取表單原值，設定初值常式：
<%=regOrgStr%>
-->
<%
'--xxxQUERY.asp--------------------------
    Set xfin = fs.opentextfile(Server.MapPath("TempQuery1.asp"))
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Query.asp")

	xfout.writeline cl & " HTProgCap=""" & request("progCapPrefix") & cq
	xfout.writeline "HTProgCode=""" & request("progRightCode") & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

	dumpTempFile
    xfin.Close

	xfout.writeline regOrgStr

    Set xfin = fs.opentextfile(Server.MapPath("TempQuery2.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline "<" & "!--#include file=""" & pgPrefix & "Form.inc""--" & ">"

    Set xfin = fs.opentextfile(Server.MapPath("TempQuery3.asp"))
	dumpTempFile
    xfin.Close
    xfout.Close

'--xxxADD.asp--------------------------
    Set xfin = fs.opentextfile(Server.MapPath("TempAdd1.asp"))
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Add.asp")

	xfout.writeline cl & " HTProgCap=""" & request("progCapPrefix") & cq
	xfout.writeline "HTProgCode=""" & request("progRightCode") & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

	dumpTempFile
    xfin.Close

	xfout.writeline "<" & "!--#include file=""" & pgPrefix & "Form.inc""--" & ">"

    Set xfin = fs.opentextfile(Server.MapPath("TempAdd2.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline ct & "sql = ""INSERT INTO " & request("tableName") & "("""

    Set xfin = fs.opentextfile(Server.MapPath("TempAdd3.asp"))
	dumpTempFile
    xfin.Close
    xfout.Close

'--xxxEDIT.asp--------------------------
    Set xfin = fs.opentextfile(Server.MapPath("TempEdit1.asp"))
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"Edit.asp")

	xfout.writeline cl & " HTProgCap=""" & request("progCapPrefix") & cq
	xfout.writeline "HTProgCode=""" & request("progRightCode") & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

	dumpTempFile
    xfin.Close

	sqlWhere = ""
	urlPara = ""
	chkPara = ""
	xcount = 0
	for each x in request.form
		if (left(x,7) = "prefix_") AND request.form(x) = "p" then
			if xcount <> 0 then		sqlWhere = sqlWhere & " & "" AND "
			if urlPara <> "" then	urlPara = urlPara & "&"
			urlPara = urlPara & mid(x,8) & "=" & cl & "=RSreg(""" & mid(x,8) & """)" & cr
			chkPara=cl & "=RSreg(""" & mid(x,8) & """)" & cr
			sqlWhere = sqlWhere & mid(x,8) & "="" & pkStr(request.queryString(""" & mid(x,8) & """),"""")"
			xcount = xcount + 1
		end if				
	next
	sqlStr = ct & "SQL = ""DELETE FROM " & request("tableName") & " WHERE " & sqlWhere
	xfout.writeline sqlStr

    Set xfin = fs.opentextfile(Server.MapPath("TempEdit2.asp"))
	dumpTempFile
    xfin.Close

	sqlStr = ct & "sqlCom = ""SELECT * FROM " & request("tableName") & " WHERE " & sqlWhere
	xfout.writeline sqlStr

    Set xfin = fs.opentextfile(Server.MapPath("TempEdit3.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline regInitStr

    Set xfin = fs.opentextfile(Server.MapPath("TempEdit4.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline "<" & "!--#include file=""" & pgPrefix & "Form.inc""--" & ">"

    Set xfin = fs.opentextfile(Server.MapPath("TempEdit5.asp"))
	dumpTempFile
    xfin.Close

	xfout.writeline ct & "sql = ""UPDATE " & request("tableName") & " SET """

    Set xfin = fs.opentextfile(Server.MapPath("TempEdit6.asp"))
	dumpTempFile
    xfin.Close
    xfout.Close

'--xxxLIST.asp--------------------------
    Set xfin = fs.opentextfile(Server.MapPath("TempList1.asp"))
    Set xfout = fs.CreateTextFile(pgPath&pgprefix&"List.asp")

	xfout.writeline cl & " HTProgCap=""" & request("progCapPrefix") & cq
	xfout.writeline "HTProgCode=""" & request("progRightCode") & cq
	xfout.writeline "HTProgPrefix=""" & pgPrefix & cq & " " & cr

	dumpTempFile
    xfin.Close

	xfout.writeline ct & "fSql = ""SELECT * FROM " & request("tableName") & " WHERE 1=1"""

    Set xfin = fs.opentextfile(Server.MapPath("TempList2.asp"))
	dumpTempFile
    xfin.Close
	for xi=1 to xItemCount
		xfout.writeline ct & "<td align=center class=lightbluetable>" & xSearchListItem(xi,0) & "</td>"
	next
    '---------------------------------------
    Set xfin = fs.opentextfile(Server.MapPath("TempList3.asp"))
	dumpTempFile
    xfin.Close
	for xi=1 to xItemCount
		if lCase(xSearchListItem(xi,2)) = "p" then
		  xfout.writeline ct & "<TD class=whitetablebg><p align=center><font size=2><a href=""" & pgPrefix & "Edit.asp?" & urlPara & """>" & cl & "=RSreg(""" & xSearchListItem(xi,1) & """)" & cr & "</A></FONT></TD>"			
		else
		  xfout.writeline ct & "<TD class=whitetablebg><p align=center><font size=2>" & cl & "=RSreg(""" & xSearchListItem(xi,1) & """)" & cr & "</FONT></TD>"			
		end if
	next
    '---------------------------
    Set xfin = fs.opentextfile(Server.MapPath("TempList4.asp"))
	dumpTempFile
    xfin.Close
    xfout.Close
    
    
sub dumpTempFile()
    Do While Not xfin.AtEndOfStream
        xinStr = xfin.readline
        xfout.writeline xinStr
    Loop
end sub
%>
</BODY>
</HTML>
