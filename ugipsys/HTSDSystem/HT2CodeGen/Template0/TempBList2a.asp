	   </select>
	   <SPAN id=RunJob style="visibility:hidden">
	   	   <input type=button class=cbutton value="�T�w����" id=button1 name=button1>
	   	   <INPUT TYPE=HIDDEN name=doJob VALUE="">
	   </SPAN>
    </td>    
    </tr>
<script language=vbs>   
      Dim chkCount
      chkCount=0            '�O��checkbox �Q�ļ�
   
      document.all.SelectJob.value="<%=fOpt%>"

    sub document_onClick           'checkbox �Q�ĭp��
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
    
     
    sub SelectJob_onChange           '�@�~���A���ܮɳB�z

		anChorURI = document.all("SelectJob").value
		xpos = inStrRev(anChorURI,".")
		if xpos>0 then	anChorURI = left(anChorURI,xpos-1)
		xPos = inStrRev(anChorURI,"/")
		if xPos>0 then
			xprogPath = left(anChorURI,xpos-1)
			anChorURI = mid(anChorURI,xPos+1)
		else
			xprogPath = "<%=request("progPath")%>"
		end if

		document.location.href= "uiView.asp?formID=" & anChorURI & "&progPath=" & xprogPath
    end sub   


   sub RunJob_Onclick                   ''�T�w���� ���Ķ�              
       reg.doJob.value = "Y"
       reg.submit
   end sub  
  
   sub Chkall
       chkCount=0     
       if document.all("ckall").value="����" then           '����
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=true
                   chkCount=chkCount+1
              end if     
          next                 
          document.all("RunJob").style.visibility="visible"
          document.all("ckall").value="������"
      elseif document.all("ckall").value="������" then        '������
          for i=0 to reg.elements.length-1
               set e=document.reg.elements(i)
               if left(e.name,5)="ckbox" then 
                   e.checked=false
               end if     
          next                 
          document.all("RunJob").style.visibility="hidden"
          document.all("ckall").value="����"          
       end if
   end sub   
</script>
  <tr>
    <td width="100%" colspan="2" class="FormRtext"></td>
  </tr>
  <tr>
    <td width="95%" colspan="2" height=230 valign=top>
  <p align="center">  
<%If not RSreg.eof then%>     
     <font size="2" color="rgb(63,142,186)"> ��
     <font size="2" color="#FF0000"><%=nowPage%>/<%=totPage%></font>��|                      
        �@<font size="2" color="#FF0000"><%=totRec%></font>
       <font size="2" color="rgb(63,142,186)">��| ���ܲ�       
         <select id=GoPage size="1" style="color:#FF0000">
             <%For iPage=1 to totPage%> 
                   <option value="<%=iPage%>"<%if iPage=cint(nowPage) then %>selected<%end if%>><%=iPage%></option>          
             <%Next%>   
         </select>      
         ��</font>           
       <% if cint(nowPage) <>1 then %>             
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage-1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�W�@��</a> 
       <%end if%>      
       
        <% if cint(nowPage)<>RSreg.PageCount and  cint(nowPage)<RSreg.PageCount  then %> 
            |<a href="<%=HTProgPrefix%>List.asp?nowPage=<%=(nowPage+1)%>&strSql=<%=strSql%>&pagesize=<%=PerPageSize%>">�U�@��</a> 
        <%end if%>     
        | �C������:
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
	<input type=button value ="����" class="cbutton"  name=ckall onClick="ChkAll"></td>
