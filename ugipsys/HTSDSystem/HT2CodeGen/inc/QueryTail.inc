<!-- µ{¦¡µ²§ô -->      
    </td>                               
  </tr> 
  <tr></tr>                              
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
      Copyright     
      &copy;2000 Chung Shan Medical and Dental College.<br>   
      Best viewed with Internet Explorer 4.0+ , Screen Size set to 800x600.</td>                                         
  </tr>                                         
</table>                                         
</body></html>        
<script language=VBS>
   for i=0 to document.all.formselect.options.length-1      
       if document.all.formselect.options(i).value="<%=HTProgPrefix%>" then      
          document.all.formselect.options(i).selected=true      
          exit for      
       end if      
   next      

  sub formselect_onChange
       if formselect.value<>"" then 
          fselect=formselect.value        
          window.parent.mainFrame.location=fselect&"Query.asp?nowKind='F'"
       end if
  end sub 
  
  ' sub document_onclick
  '      set sObj=window.event.srcElement
  '      if  sObj.tagname="INPUT" and left(sObj.name,5)="PABDB" then
  '          window.navigate sObj.name&"Query.asp"
  '     end if'
'end sub
</script>      
</script>                                 