<%

  'if oxml.parseError.reason <> "" then 
  '  Response.Write("xhtPageDom parseError on line " &  oxml.parseError.line)
  '  Response.Write("<BR>Reason: " &  oxml.parseError.reason)
  '  Response.End()
  'end if


  if (session("ErrCount_" & Request.ServerVariables("url")) = "") or oxml.parseError.reason = "" then
    session("ErrCount_" & Request.ServerVariables("url")) = 0
  end if
  if oxml.parseError.reason <> "" then 
    if cint(session("ErrCount_" & Request.ServerVariables("url"))) > 3 then        
        Response.Write("htPageDom parseError on line " &  oxml.parseError.line)
        Response.Write("<BR>Reason: " &  oxml.parseError.reason)
        Response.Write("<BR>" &  session("ErrCount_" & Request.ServerVariables("url")))
        session("ErrCount_" & Request.ServerVariables("url")) = 0
        Response.End()
    else
        %>
        <script type="text/javascript">window.location.reload()</script>
        <%
        session("ErrCount_" & Request.ServerVariables("url")) = cint(session("ErrCount_" & Request.ServerVariables("url"))) + 1
        response.End()
    end if  
  end if
%>
