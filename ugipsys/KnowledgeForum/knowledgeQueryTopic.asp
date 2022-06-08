<%Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="GC1AP9"
HTProgPrefix="newkpi" %>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="/css/list.css" rel="stylesheet" type="text/css">
<link href="/css/layout.css" rel="stylesheet" type="text/css">
<title></title>
</head>
<%
    if request("reset")="Y" then
        session("sTitle") = ""
        session("xbody") = ""
    end if
%>
<Form id="reg" name="reg" method="POST" action="KnowledgeForumlist.asp">
 <INPUT TYPE=hidden name=submitTask value="">
 <INPUT type='hidden' id='searchType' name='searchType' value="">
 
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
	  <td width="50%" align="left" nowrap class="FormName">問題標題查詢&nbsp;</td>
		<td width="50%" class="FormLink" align="right">
			<A href="Javascript:window.location.href='KnowledgeForumList.asp';" title="回前頁">回前頁</A>			
		</td>	
  </tr>
	<tr>
	  <td width="100%" colspan="2"><hr noshade size="1" color="#000080"></td>
	</tr>
	<tr>
		<td class="Formtext" colspan="2" height="15"></td>
	</tr>  
	<tr>
		<td align=center colspan=2 width=80% height=230 valign=top>    
		<object data="/inc/calendar.htm" id=calendar1 type=text/x-scriptlet width=245 height=160 style='position:absolute;top:0;left:230;visibility:hidden'></object>
		<INPUT TYPE="hidden" name=CalendarTarget>
		<CENTER>
		<TABLE border="0" id="ListTable" cellspacing="1" cellpadding="2">
		<TR>
			<Th>發佈日期</Th>
			<TD class="eTableContent">
				<input name="xPostDateS" size="10" readonly onclick="VBS: popCalendar 'xPostDateS','xPostDateE'"> ～ 				
				<input name="xPostDateE" size="10" readonly onclick="VBS: popCalendar 'xPostDateE',''">
			</TD>
		</TR>
		<TR>
			<Th>討論關閉</Th>
			<TD class="eTableContent">
				<select name="xNewWindow" id="xNewWindow" >
  		
    <option value="" <%if xNewWindow="" then%>selected<%END IF%>>請選擇</option>
    <option value="Y" <%if xNewWindow="Y" then%>selected<%END IF%>>是</option>
    <option value="N" <%if xNewWindow="N" then%>selected<%END IF%>>否</option>
   
  </select>
			</TD>
		</TR>
		<TR>
			<Th>狀態</Th>
			<TD class="eTableContent">
				<select name="Status" id="Status" >
  		
    <option value="" <%if Status="" then%>selected<%END IF%>>請選擇</option>
    <option value="N" <%if Status="N" then%>selected<%END IF%>>正常</option>
    <option value="D" <%if Status="D" then%>selected<%END IF%>>刪除</option>
    
  </select>
			</TD>
		</TR>
		<TR>
			<Th>是否公開</Th>
			<TD class="eTableContent">
				<select name="fCTUPublic" id="fCTUPublic" >
  		
    <option value="" <%if fCTUPublic="" then%>selected<%END IF%>>請選擇</option>
    <option value="Y" <%if fCTUPublic="Y" then%>selected<%END IF%>>公開</option>
    <option value="N" <%if fCTUPublic="N" then%>selected<%END IF%>>不公開</option>
  </select>
			</TD>
		</TR>
		<TR>
			<Th>是否為知識家活動問題</Th>
			<TD class="eTableContent">
				<select name="vGroup" id="vGroup" >
  		
    <option value="" <%if vGroup="" then%>selected<%END IF%>>請選擇</option>
    <option value="A" <%if vGroup="A" then%>selected<%END IF%>>是</option>
    <option value="N" <%if vGroup <> "" then%>selected<%END IF%>>否</option>
	<option value="OFF" <%if vGroup = "OFF" then%>selected<%END IF%>>已下架</option>
  </select>
			</TD>
		</TR>
		<TR>
		<Th>上傳者(帳號、姓名、暱稱)</Th>
			<TD class="eTableContent"><input name="memberId" size="50" value="" ></TD>
		</TR>
			<TR>
			<Th>問題標題</Th>
			<TD class="eTableContent"><input name="sTitle" size="50" value="<%=session("sTitle") %> "></TD>
		</TR>
		<TR>
			<Th>內文</Th>
			<TD class="eTableContent"><input name="xbody" size="50" value="<%=session("xbody") %> "></TD>
		</TR>
		
		</TABLE>
		</td>
	</tr>
	</table>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td width="100%">     
			<p align="center">
				<input name="button" type="button" class="cbutton" onclick="QuerySubmit('N')" value ="查 詢">
				<input name="button" type="button" class="cbutton" onclick="QuerySubmit('R')" value ="再次查詢">
                <input type="button" value ="重　填" class="cbutton" onClick="returnToEdit()">        
		</td>
	</tr>
	</table>

</TABLE>
         
       <!-- 程式結束 ---------------------------------->  
</form>  
</body>
</html>     

<script language="javascript">
	function QuerySubmit(type) 
	{
	    //document.getElementById("picId").value = id;
		//document.getElementById("parentIcuitem").value = parentIcuitem;

	    document.getElementById("searchType").value = type;	    
		document.getElementById("reg").submit();
	}
	function returnToEdit() {
		//window.location.href = "cp_question.asp?iCUItem=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=cint(request.querystring("pagesize"))%>";
	     window.location.href = "knowledgeQueryTopic.asp?reset=Y";
	
	}
	
</script>    
<script language=vbs>
	cvbCRLF = vbCRLF
	cTabchar = chr(9)

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
			'document.all("x"&CanTarget).value=document.all.CalendarTarget.value
			if followCanTarget<>"" then
				document.all(followCanTarget).value=document.all.CalendarTarget.value
				'document.all("x"&followCanTarget).value=document.all.CalendarTarget.value
			end if
		end if
	end sub   
 	
	
</script> 
                                                   

</body>
</html>
