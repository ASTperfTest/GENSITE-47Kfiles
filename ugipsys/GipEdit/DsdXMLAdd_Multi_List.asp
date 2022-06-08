<%@ CodePage = 65001 %>
<%
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="多向出版"
HTProgFunc="參照清單"
HTProgCode="GC1AP1"
HTProgPrefix="DsdXMLAdd_Multi" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
Dim RSreg
if request("submitTask")<>"" then
    SQLDelete=""
    For Each x In Request.Form
	if left(x,5)="ckbox" and request(x)<>"" then
	    xn=mid(x,6)
	    SQLDelete=SQLDelete+"Delete CuDtGeneric where icuitem=" & request("xphKeyiCUItem"&xn) & ";"
        end if
    Next 
    if SQLDelete<>"" then conn.execute(SQLDelete)
%>
	<script language=VBS>
		alert "刪除完成！"
		window.navigate "DsdXMLEdit.asp?iCuItem=<%=request("iCuItem")%>&phase=edit"
	    </script>
<%	   
	response.end
else
    	fSql=session("DsdXMLAdd_Multi_List")
	set RSreg=conn.execute(fSql)
end if
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
<script Language=VBScript>

      Dim chkCount
      chkCount=0            '記錄checkbox 被勾數

    sub document_onClick           'checkbox 被勾計數
         set sObj=window.event.srcElement
         if sObj.tagName="INPUT" then 
            if sObj.type="checkbox"  then 
                if sObj.checked then 
                   chkCount=chkCount+1
                else
                   chkCount=chkCount-1                
                end if                                          
            end if
         end if
         '
         if chkCount=0 then 
            document.all("RunJob").style.visibility="hidden"
         else
            document.all("RunJob").style.visibility="visible"
         end if
    end sub        
    sub Chkall
       chkCount=0     
       if document.all("ckall").value="全選" then           '全勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=true
                   chkCount=chkCount+1
              end if     
          next                 
          document.all("RunJob").style.visibility="visible"
          document.all("ckall").value="全不選"
      elseif document.all("ckall").value="全不選" then        '全不勾
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=false
               end if     
          next                 
          document.all("RunJob").style.visibility="hidden"
          document.all("ckall").value="全選"          
       end if
   end sub   
   sub button1_onClick
        chky=msgbox("注意！"& vbcrlf & vbcrlf &"　您確定要刪除資料嗎？"& vbcrlf , 48+1, "請注意！！")
        if chky=vbok then
	      reg.submitTask.value = "UPDATE"
	      reg.Submit
       	end If      
   end sub

</script>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
 <Form name=reg method="POST" action=<%=HTprogPrefix%>_List.asp>
 <INPUT TYPE=hidden name=submitTask value="">
 <input name=iCuItem type="hidden" value="<%=request.querystring("iCuItem")%>">
  <tr>
	    <td class="FormName"><%=HTProgCap%>&nbsp;
	    <font size=2>【<%=HTProgFunc%>】</td>
    	<td class="FormLink" valign="top" align=right>
		<A href="DsdXMLAdd_Multi.asp?<%=request.querystring%>" title="回前頁">回前頁</A>
    	</td>    	    
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top" align=right colspan="2">
	   <SPAN id=RunJob style="visibility:hidden">
	   	   <input type=button class=cbutton value="刪除" id=button1 name=button1>
	   </SPAN>
    </td>    
    </tr>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
	<td align=center class=eTableLable width=7%>刪除<br>
	<input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></td>
     	<td class=eTableLable>單元名稱</td>
     	<td class=eTableLable>標  題</td>
     	<td class=eTableLable>部  門</td> 
     	<td class=eTableLable>編修日期</td> 
     	<td class=eTableLable>編修</td>      	
     </tr> 	     	     	 
<%
If not RSreg.eof then   

    while not RSreg.eof 
      	i = i+1
%>                  
<tr>                  
<TD align=center><INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y">
<INPUT TYPE=hidden name="xphKeyiCUItem<%=i%>" value="<%=RSreg("iCUItem")%>"></td>   
	<TD class=eTableContent><font size=2>
	<%=RSreg("CatName")%></TD>	
	<TD class=eTableContent><font size=2><%=RSreg("sTitle")%>
	</TD>
	<TD class=eTableContent><font size=2>
	<%=RSreg("DeptName")%></TD>			
	<TD class=eTableContent><font size=2>
	<%=RSreg("dEditDate")%></TD>
	<td align=center class=eTableContent>
	<input type=button value ="編修" class="cbutton" onClick="DsdXMLEdit '<%=RSreg("iCUItem")%>'"></td>				
    </tr>
    <%
         RSreg.moveNext
    wend
   %>
<%else%>
    <tr>                  
    <TD align=center class=eTableContent colspan=7><font size=2>無參照資料!</font>
    </td></tr>
<%end if%>

    </TABLE>    
    </CENTER>
     </form>      
    </td>
  </tr>    
</table>   
</body>
</html>  
<script language=VBS>
sub DsdXMLEdit(iCuItem)
'document.location.reload
	window.open "DsdXMLEdit.asp?phase=edit&S=Approve&iCuItem="&iCuItem,"","scrollbars=yes,width=800,height=600"
end sub
sub GipApproveListRefresh()
	document.location.reload
end sub
</script>


