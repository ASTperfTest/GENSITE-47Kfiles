<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料審稿"
HTProgFunc="查詢"
HTProgCode="GC1AP2"
HTProgPrefix="GipApprove" %>
<!--#include virtual = "/inc/server.inc" -->
    <html> 
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
    <title>查詢表單</title>
    <link rel="stylesheet" type="text/css" href="/inc/setstyle.css">
    </head>
<!--#include virtual = "/inc/dbutil.inc" -->
<!--#include virtual = "/inc/selectTree.inc" -->
<%
taskLable="查詢" & HTProgCap

	showHTMLHead()
	formFunction = "query"
	showForm()
	initForm()
	showHTMLTail()
%>

<% sub initForm() %>
<script language=vbs>
sub resetForm
	reg.reset()
	clientInitForm
end sub

sub clientInitForm

    end sub

    sub window_onLoad  
	     clientInitForm
    end sub

Dim CanTarget
Dim followCanTarget

sub popCalendar(dateName,followName)        
 	CanTarget=dateName
 	followCanTarget=followName
	xdate = document.all(CanTarget).value
	if not isDate(xdate) then	xdate = date()
	document.all.calendar1.setDate xdate
	
 	If document.all.calendar1.style.visibility="" Then           
   		document.all.calendar1.style.visibility="hidden"        
 	Else        
       ex=window.event.clientX
       ey=document.body.scrolltop+window.event.clientY+10
       if ex>520 then ex=520
       document.all.calendar1.style.pixelleft=ex-80
       document.all.calendar1.style.pixeltop=ey
       document.all.calendar1.style.visibility=""
 	End If              
end sub     

sub calendar1_onscriptletevent(n,o) 
    	document.all("CalendarTarget").value=n         
    	document.all.calendar1.style.visibility="hidden"    
    if n <> "Cancle" then
        document.all(CanTarget).value=document.all.CalendarTarget.value
        document.all("pcShow"&CanTarget).value=d7date(document.all.CalendarTarget.value)
        if followCanTarget<>"" then
        	document.all(followCanTarget).value=document.all.CalendarTarget.value
        	document.all("pcShow"&followCanTarget).value=d7date(document.all.CalendarTarget.value)
        end if
	end if
end sub       
  </script>	
<%end sub '---- initForm() ----%>

<%Sub showForm() %>
    <script language="vbscript">
    sub formSearchSubmit()
         reg.action="<%=HTprogPrefix%>List.asp"
         reg.Submit
    end Sub
   </script>

<form method="POST" name="reg" action="GipApproveList.asp">
<INPUT TYPE=hidden name=submitTask value="">
<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
<INPUT TYPE=hidden name=CalendarTarget>
<CENTER><TABLE border="0" class="eTable" cellspacing="1" cellpadding="2" width="90%">
<TR>
  <TD class="eTableLable" align="right">目錄樹：</TD>
  <TD class="eTableContent"><SELECT name="htx_ctRootID" size=1>
	 	<option value="">--不 拘--</option>
<%
	SQL = "" & _
		" Select ctRootID, ctRootName from CatTreeRoot where inUse = N'Y' AND pvXdmp IS NOT NULL " & _
		" and (deptID is null or deptID like '" & session("deptID") & "%') order by deptID "
		SET RSS=conn.execute(SQL)
		While not RSS.EOF%>
	
		<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
		<%	RSS.movenext
		wend%>
  </SELECT></TD>
</TR>
<TR><TD class="eTableLable" align="right">單元名稱：</TD>
<TD class="eTableContent"><input name="htx_CatName" size="30">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">標題：</TD>
<TD class="eTableContent"><input name="htx_sTitle" size="30">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">內容：</TD>
<TD class="eTableContent"><input name="htx_xBody" size="30">
</TD>
</TR>
<TR><TD class="eTableLable" align="right">狀態：</TD>
<TD class="eTableContent"><Select name="htx_fctupublic" size=1>
         	<option value="">--所有待審--</option>
 			<%SQL="Select mcode,mvalue from CodeMain where codeMetaId=N'isPublic3' Order by msortValue DESC"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
        	
         </select>
    </td>    
  </tr>
<TR><TD align="right" class="eTableLable">單 位：</TD>
<TD class="eTableContent"><Select name="htx_deptId" size=1>
<%
	sqlCom ="SELECT D.deptId,D.deptName,D.parent,len(D.deptId)-1,D.seq," & _
		"(Select Count(*) from Dept WHERE parent=D.deptId AND nodeKind='D') " & _
		"  FROM Dept AS D Where D.nodeKind='D' " _
		& " AND D.deptId LIKE N'" & session("deptId") & "%'" _
		& " ORDER BY len(D.deptId), D.parent, D.seq"	
'response.write SqlCom
	set RSS = conn.execute(sqlCom)
	if not RSS.EOF then
		ARYDept = RSS.getrows(300)
		glastmsglevel = 0
		genlist 0, 0, 1, 0
	        expandfrom ARYDept(cid, 0), 0, 0
	end if
%>
</select>
</TD></TR>
  <tr>
    <td class="eTableLable" align="right">編修日期</td>
    <td class="eTableContent" colspan=3>
   
    <input name="pcShowhtx_IDateS" type="text" class="InputText" size="10" readonly onclick="VBS: popCalendar 'htx_IDateS','htx_IDateE'">
      ～<input name="htx_IDateS" type=hidden><input name="htx_IDateE" type=hidden>
        <input name="pcShowhtx_IDateE" type="text" class="InputText" size="10" readonly onclick="VBS: popCalendar 'htx_IDateE',''">
        </td>
  </tr>
</TABLE>
</CENTER>
</form>     
<table border="0" width="100%" cellspacing="0" cellpadding="0">
 <tr><td width="100%">     
   <p align="center">
<%If formFunction = "query" then %>
        <%if (HTProgRight AND 2) <> 0 then %>
            <input type=button value ="查　詢" class="cbutton" onClick="formSearchSubmit()">
            <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <%end if%>    
<%Elseif formFunction = "edit" then %>
        <%if (HTProgRight AND 8) <> 0 then %>
               <input type=button value ="編修存檔" class="cbutton" onClick="formModSubmit()">
        <% End If %>
        
       <% if (HTProgRight AND 16) <> 0 then %>
             <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
       <%end If %>           
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">

<% Else '-- add ---%>          
          <%if (HTProgRight AND 4)<>0 then %>
               <input type=button value ="新增存檔" class="cbutton" onClick="formAddSubmit()">
          <%end if%>     
          
          <%if (HTProgRight AND 4)<>0  then %>
              <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
          <%end if%>    

<% End If %>
 </td></tr>
</table>
<SCRIPT LANGUAGE="vbs">
    sub setOtherCheckbox(xname,otherValue)
    	reg.all(xname).item(reg.all(xname).length-1).value = otherValue
    end sub
</SCRIPT>
	
<%end sub '--- showForm() ------%>

<%sub showHTMLHead() %>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
			<A href="Javascript:window.history.back();">回前頁</A> 
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
<%end sub '--- showHTMLHead() ------%>

<%sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<%end sub '--- showHTMLTail() ------%>
