<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="資料審稿"
HTProgFunc="清單"
HTProgCode="GC1AP2"
HTProgPrefix="CuCheck" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%

Set RSreg = Server.CreateObject("ADODB.RecordSet")
nowPage=Request.QueryString("nowPage")  '現在頁數

if request("doJob")<>"" then
    SQLUpdate=""
    For Each x In Request.Form
	if left(x,5)="ckbox" and request(x)<>"" then
	    xn=mid(x,6)
	    SQLUpdate=SQLUpdate+"Update xDictonary Set xEngWord=N'"&request("htx_xEngWord"&xn)&"', xChnWord=N'"&request("htx_xChnWord"&xn)&"',wordKind=N'"&request("htx_wordKind"&xn)&"', " & _
	    	"wordStatus='Y' where xDicIID=" & request("xDicIID"&xn) & ";"
        end if
    Next
    if SQLUpdate<>"" then conn.execute(SQLUpdate)
%>
	<script language=VBS>
		alert "資料審稿完成完成！"
		window.navigate "<%=HTProgPrefix%>List.asp?nowPage=1"
	    </script>
<%	    response.end
else
    if nowPage="" then
	fSql = "SELECT htx.*, xref1.mValue AS xref1wordKind, xref2.mValue AS xref2wordStatus" _
		& " FROM ((xDictonary AS htx LEFT JOIN codeMain AS xref1 ON xref1.mCode = htx.wordKind AND xref1.codeMetaID='wordKind') LEFT JOIN codeMain AS xref2 ON xref2.mCode = htx.wordStatus AND xref2.codeMetaID='wordState')" _
		& " WHERE 1=1 "
	xpCondition
	session("CuCheckList")=fSql
    else
	fSql=session("CuCheckList")
    end if
end if


'----------HyWeb GIP DB CONNECTION PATCH----------
'RSreg.Open fSql,Conn,1,1
Set RSreg = Conn.execute(fSql)

'----------HyWeb GIP DB CONNECTION PATCH----------


if Not RSreg.eof then 
   totRec=RSreg.Recordcount       '總筆數
   if totRec>0 then 
      PerPageSize=cint(Request.QueryString("pagesize"))
      if PerPageSize <= 0 then  
         PerPageSize=10  
      end if 
      
      RSreg.PageSize=PerPageSize       '每頁筆數

      if cint(nowPage)<1 then 
         nowPage=1
      elseif cint(nowPage) > RSreg.PageCount then 
         nowPage=RSreg.PageCount 
      end if            	

      RSreg.AbsolutePage=nowPage
      totPage=RSreg.PageCount       '總頁數
   end if    
end if   

function qqRS(fldName)
	if request("submitTask")="" then
		xValue = RSreg(fldname)
	else
		xValue = ""
		if request("htx_"&fldName) <> "" then
			xValue = request("htx_"&fldName)
		end if
	end if
	if isNull(xValue) OR xValue = "" then
		qqRS = ""
	else
		xqqRS = replace(xValue,chr(34),chr(34)&"&chr(34)&"&chr(34))
		qqRS = replace(xqqRS,vbCRLF,chr(34)&"&vbCRLF&"&chr(34))
	end if
end function  
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="/inc/setstyle.css">
<title></title>
</head>
<body>
 <Form name=reg method="POST" action=<%=HTprogPrefix%>List.asp>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
		<%if (HTProgRight and 1)=1 then%>
			<A href="CuCheckQuery.asp" title="重設查詢條件">查詢</A>
		<%end if%>
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15" align=right>
	   <SPAN id=RunJob style="visibility:hidden">
	   	   <input type=button class=cbutton value="審核通過" id=button1 name=button1>
	   	   <INPUT TYPE=HIDDEN name=doJob VALUE="">
	   </SPAN></td>	    
	  </tr>  
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
  <p align="center">  
<%If not RSreg.eof then%>     
     <font size="2" color="rgb(63,142,186)"> 第
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>頁|                      
        共<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">筆| 跳至第       
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
         頁</font>           
       <% if cint(nowPage) <>1 then %>             
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&pagesize=<%=PerPageSize%>">上一頁</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&pagesize=<%=PerPageSize%>">下一頁</a> 
        <%end if%>     
        | 每頁筆數:
       <select id=PerPage size="1" style="color:#FF0000">            
             <option value="10"<%if PerPageSize=10 then%> selected<%end if%>>10</option>                       
             <option value="20"<%if PerPageSize=20 then%> selected<%end if%>>20</option>
             <option value="30"<%if PerPageSize=30 then%> selected<%end if%>>30</option>
             <option value="50"<%if PerPageSize=50 then%> selected<%end if%>>50</option>
        </select>     
     </font>     
    <CENTER>
     <TABLE width=100% cellspacing="1" cellpadding="0" class=bg>                   
     <tr align=left>    
	<td align=center class=eTableLable width=7%>
	<input type=button value ="全選" class="cbutton"  name=ckall onClick="ChkAll"></td>    
	<td class=eTableLable>編修日期</td>	     
	<td class=eTableLable>英文</td>
	<td class=eTableLable>中文</td>
	<td class=eTableLable>詞性</td>
	<td class=eTableLable>狀態</td>
	<td class=eTableLable>點選計數</td>
	<td class=eTableLable>權重</td>
    </tr>	                                
<%                   
    for i=1 to PerPageSize                  
%> 
<tr>       
<TD align=center>
<%if RSreg("wordStatus")="N" or RSreg("wordStatus")="P" then%>
	<INPUT TYPE=checkbox name="ckbox<%=i%>" value="Y">
	<INPUT TYPE=hidden name="xDicIID<%=i%>" value="<%=RSreg("xDicIID")%>">         
<%else%>
	&nbsp;
<%end if%>
</td>                               
<%pKey = ""
pKey = pKey & "&xDicIID=" & RSreg("xDicIID")
if pKey<>"" then  pKey = mid(pKey,2)%>
	<TD class=eTableContent><font size=2>
<A href="xdic2Edit.asp?<%=pKey%>"><%=d7date(RSreg("EditDate"))%>
</font></td>
	<TD class=eTableContent><font size=2>
<input name="htx_xEngWord<%=i%>" size="35">
<script language=vbs>
	reg.all("htx_xEngWord<%=i%>").value= "<%=qqRS("xEngWord")%>"
</script>
</font></td>
	<TD class=eTableContent><font size=2>
<input name="htx_xChnWord<%=i%>" size="35">
<script language=vbs>
	reg.all("htx_xChnWord<%=i%>").value= "<%=qqRS("xChnWord")%>"
</script>
</font></td>
	<TD class=eTableContent><font size=2>
<Select name="htx_wordKind<%=i%>" size=1>
<option value="">請選擇</option>
			<%SQL="Select mCode,mValue from codeMain where mSortValue IS NOT NULL AND codeMetaID='wordKind' Order by mSortValue"
			SET RSS=conn.execute(SQL)
			While not RSS.EOF%>

			<option value="<%=RSS(0)%>"><%=RSS(1)%></option>
			<%	RSS.movenext
			wend%>
</select>

<script language=vbs>
	reg.all("htx_wordKind<%=i%>").value= "<%=qqRS("wordKind")%>"
</script>
</font></td>
	<TD class=eTableContent><font size=2>
<%=RSreg("xref2wordStatus")%>
</font></td>
	<TD class=eTableContent align="center"><font size=2>
<%=RSreg("clkCount")%>
</font></td>
	<TD class=eTableContent align="right"><font size=2>
<%=RSreg("evScore")%>
</font></td>
    </tr>
    <%
         RSreg.moveNext
         if RSreg.eof then exit for 
    next 
   %>
    </TABLE>
    </CENTER>
       <!-- 程式結束 ---------------------------------->  
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 
</form>
</body>
</html>                                 
<%else%>
      <script language=vbs>
           msgbox "找不到資料, 請重設查詢條件!"
'	       window.history.back
      </script>
<%end if%>

<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=<%=PerPageSize%>"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="<%=HTProgPrefix%>List.asp?nowPage=<%=nowPage%>" & "&strSql=<%=strSql%>&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

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
   sub button1_Onclick                   ''確定執行 打勾項              
       reg.doJob.value = "Y"
       reg.submit
   end sub 

</script>
