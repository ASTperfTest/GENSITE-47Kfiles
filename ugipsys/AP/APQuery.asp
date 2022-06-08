<%@ CodePage = 65001 %>
<% Response.Expires = 0 
HTProgCap="AP"
HTProgCode="HT003"
HTProgPrefix="AP" %>
<!--#include virtual = "/inc/server.inc" -->
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
	reg.pfx_APcode.value = "<%=request("reg.pfx_APcode")%>"
	reg.tfx_APnameE.value = "<%=request("reg.tfx_APnameE")%>"
	reg.tfx_APnameC.value = "<%=request("reg.tfx_APnameC")%>"
	reg.sfx_Apcat.value = "<%=request("reg.sfx_Apcat")%>"
	reg.tfx_APpath.value = "<%=request("reg.tfx_APpath")%>"
	reg.tfx_APorder.value = "<%=request("reg.tfx_APorder")%>"
	reg.nfx_APMask.value = "<%=request("reg.nfx_APMask")%>"
	reg.tfx_Spare64.value = "<%=request("reg.tfx_Spare64")%>"
	reg.tfx_Spare128.value = "<%=request("reg.tfx_Spare128")%>"

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

<!--#include file="APForm.inc"-->
<%end sub '--- showForm() ------%>

<%sub showHTMLHead() %>
    <html> 
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
    <meta name="GENERATOR" content="Hometown Code Generator 1.0">
    <meta name="ProgId" content="<%=HTprogPrefix%>Query.asp">
    <title>查詢表單</title>
    <link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
    </head>
    <body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=Title%>&nbsp;<font size=2>【<%=HTProgCap%>查詢引擎】</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <%if (HTProgRight and 4)=4 then%>
	       <a href="<%=HTprogPrefix%>Add.asp">新增</a>	    
	       <%end if%>
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
  <tr>                               
    <td width="100%" colspan="2" align="center" class="English">                               
	<!--#include virtual = "/inc/Footer.inc"-->
      </td>                                         
  </tr> 
</table> 
</body>
</html>
<%end sub '--- showHTMLTail() ------%>
