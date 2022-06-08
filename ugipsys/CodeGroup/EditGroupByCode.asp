<%@ CodePage = 65001 %>
<% Response.Expires = 0 
   HTProgCode = "Pn50M06"  %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Function TranComp(groupcode,value)
   if (groupcode and value )=value then
      response.write "checked"
   end if
End function

Function TranComp1(groupcode,value)
   if (groupcode and value ) <> value then
      response.write "none"
   end if
End function

If Request.Form("Enter") = "確定存檔" Then   
     Set conn2 = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'     conn2.Open session("ODBCDSN")  
'Set conn2 = Server.CreateObject("HyWebDB3.dbExecute")
conn2.ConnectionString = session("ODBCDSN")
conn2.ConnectionTimeout=0
conn2.CursorLocation = 3
conn2.open
'----------HyWeb GIP DB CONNECTION PATCH----------

     CodeID= pkstr(Request("CodeID"),"")
     SQLDelete=""
     SQLInsert=""
     DataDate = Date()
     SysUser = Session("UserID")
     SQLDelete = "Delete from ugrpCode Where CodeID = " & CodeID
     SQLCom = "SELECT ugrpID FROM ugrp"
     set RSCom = Conn.execute(sqlcom)
      while not RSCom.eof
           sumvalue=0
           for i=1 to 8
               ugrpIDName=RSCom("ugrpID")&i
               if request.Form(ugrpIDName)="Y" then 
                    if i=1 then
                       sumvalue=sumvalue+1
                    elseif i=2 then
                       sumvalue=sumvalue+2
                    elseif i=3 then
                       sumvalue=sumvalue+4
                    elseif i=4 then
                       sumvalue=sumvalue+8
                    elseif i=5 then
                       sumvalue=sumvalue+16
                    elseif i=6 then
                       sumvalue=sumvalue+32
                    elseif i=7 then
                       sumvalue=sumvalue+64
                    elseif i=8 then
                       sumvalue=sumvalue+128                      
                    end if
               end if
           next
           mStart = request.Form(RSCom("ugrpID") & "start")
           mEnd = request.Form(RSCom("ugrpID") & "end")
           if trim(mEnd)="" then
			   mEnd="2079/6/6"		' SmallDate 最大值
	   end if
	   if sumvalue<>0 then		   		   
           	SQLInsert = SQLInsert & "INSERT INTO ugrpCode (ugrpID, CodeID, Rights,regdate,startdate,enddate) VALUES (N'"& RSCom("ugrpID") &"', "& CodeID &", "& sumvalue &",N'" & date() & "',N'" & mStart & "',N'" & mEnd & "');"           
	   end if
     RSCom.movenext
     wend
'response.write SQLDelete & "<hr>"     
'response.write SQLInsert & "<hr>"   
'response.end     
     conn2.BeginTrans 
     	conn2.execute(SQLDelete)
     	if SQLInsert<>"" then     	
     		conn2.execute(SQLInsert)
     	end if
	if err.Number<>0 then
		conn2.RollBack	
	else
		conn2.CommitTrans
	end if	       
     %>
		<script language="VBScript">
		  alert("編修完成！")
		  window.navigate "EditGroupByGroup.asp?ugrpID=<%=request("ugrpID")%>"
		</script>	
<%     		response.end
End IF
 
	CodeID = pkstr(Request.querystring("CodeID"),"")
	SQL="Select CodeType,CodeName from CodeMetaDef where CodeID=" & CodeID
	Set RSU=conn.execute(SQL)
	SQLCom = "SELECT U.UgrpID, U.ugrpName ,UA.rights URights, UA.StartDate, UA.EndDate FROM ugrp U " & _
			"Left Join ugrpCode UA ON U.ugrpID=UA.ugrpID AND UA.CodeID= " & CodeID & _
			"Order by U.ugrpID"		
	Set RS = Conn.execute(SQLCom)	
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>編修代碼權限群組(by 代碼)</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
    <td width="100%" class="FormName" colspan="2"><%=Title%></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="Formtext" width="60%">【編修代碼權限群組(by 代碼)】</td>
    <td class="FormLink" valign="top" width="40%">
      <p align="right"><% if (HTProgRight and 2)=2 then %><a href="Querygroup.asp">查詢群組</a>　<% End IF %><a href="VBScript: history.back">回上一頁</a>&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td align="right" class="whitetablebg">[ <font color="red"><%=RSU("CodeType")%>--<%=RSU("CodeName")%></font> ]
    </td>
    <td>

    </td>
  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>  
  <tr>
    <td width="100%" colspan="2" id="rightsTable">
    <form method="POST" action="EditGroupByCode.asp" name="reg">
    <input type="hidden" name="CodeID" value="<%=request.querystring("CodeID")%>"> 
    <input type="hidden" name="ugrpID" value="<%=request.querystring("ugrpID")%>">     
    <% If Not rs.Eof Then %>
        <center>
        <table border="0" width="100%" cellspacing="1" class="bluetable" style="Display:'';" ID="M<%=cno%>" cellpadding="2">
          <tr class="lightbluetable">
            <td width="5%" align="center"></td>
            <td width="30%" align="center">群組名稱</td>
            <td width="7%" align="center">查詢</td>
            <td width="7%" align="center">新增</td>
            <td width="7%" align="center">修改</td>
            <td width="7%" align="center">刪除</td>
            <td width="7%" align="center">列印</td>
            <td width="7%" align="center">備用A</td>            
            <td width="7%" align="center">備用B</td>
            <td width="8%" align="center">起始日期時間</td>            
            <td width="8%" align="center">終止日期時間</td>
          </tr>
           <% do while not rs.eof %>
          <tr>
            <td width="5%" align="center" class="lightbluetable">
             <input type="checkbox" name="<%=rs("ugrpID")%>1" value="Y" OnClick="VBScript: chkbox '<%=rs("ugrpID")%>'" <%=TranComp(RS("URights"),1)%>></td>
            <td width="32%" class="whitetablebg"><%=rs("ugrpName")%></A></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("ugrpID")%>2" value="Y" style="display:'<%=TranComp1(RS("URights"),1)%>'" id="x<%=rs("ugrpID")%>2" <%=TranComp(RS("URights"),2)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("ugrpID")%>3" value="Y" style="display:'<%=TranComp1(RS("URights"),1)%>'" id="x<%=rs("ugrpID")%>3" <%=TranComp(RS("URights"),4)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("ugrpID")%>4" value="Y" style="display:'<%=TranComp1(RS("URights"),1)%>'" id="x<%=rs("ugrpID")%>4" <%=TranComp(RS("URights"),8)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("ugrpID")%>5" value="Y" style="display:'<%=TranComp1(RS("URights"),1)%>'" id="x<%=rs("ugrpID")%>5" <%=TranComp(RS("URights"),16)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("ugrpID")%>6" value="Y" style="display:'<%=TranComp1(RS("URights"),1)%>'" id="x<%=rs("ugrpID")%>6" <%=TranComp(RS("URights"),32)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("ugrpID")%>7" value="Y" style="display:'<%=TranComp1(RS("URights"),1)%>'" id="x<%=rs("ugrpID")%>7" <%=TranComp(RS("URights"),64)%>></td>
            <td width="7%" align="center" class="whitetablebg">
             <input type="checkbox" name="<%=rs("ugrpID")%>8" value="Y" style="display:'<%=TranComp1(RS("URights"),1)%>'" id="x<%=rs("ugrpID")%>8" <%=TranComp(RS("URights"),128)%>></td>
            <td width="7%" class="whitetablebg"><input type="text" name="<%=rs("ugrpID")%>start" value="<%=rs("startdate")%>" size="12"></td>
            <td width="7%" class="whitetablebg"><input type="text" name="<%=rs("ugrpID")%>end" value="<%=rs("enddate")%>" size="12"></td>
          </tr>
          <%  rs.movenext      
              loop %>
        </table>
        </center>
       <%End If %>
    <p align="center">  
    <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3"><br>
    <% if (HTProgRight and 8)=8 then %><input type="submit" value="確定存檔" Name="Enter" class="cbutton"><% End If %>
    </form>
    </td>
  </tr>
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      <!--#include virtual = "/inc/Footer.inc" -->
      </td>                                         
  </tr>                    
</table>               
</body></html>               
<script language="VBScript">      
Sub chkbox(x)     
	if reg(x & "1").checked then     
  		for i=2 to 8     
    			document.all("x" & x & i).style.display=""         
  		next     
 	Else     
  		for i=2 to 8    
    			document.all("x" & x & i).style.display="none"         
    			reg(x & i).checked=false     
  		next     
 	end if     
end sub   
  
sub rightsTable_onDblClick     
	set xObj = window.event.srcElement    
	if xObj.tagname="INPUT" then 
	if xObj.type = "checkbox" then     
		xname=xObj.name     
		xnPos = cint(right(xname,1))     
		xname = left(xname, len(xname)-1)     
		xBoolean = xObj.checked     
		for i=2 to xnPos     
			document.all(xname&i).checked = xBoolean     
		next     
	end if     
	end if
end sub     
</script>         
