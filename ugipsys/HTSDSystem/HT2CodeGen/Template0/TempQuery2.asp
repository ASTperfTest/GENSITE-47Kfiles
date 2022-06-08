    end sub

    sub window_onLoad  
	     clientInitForm
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

