                   
<% end sub '--- showForm() ------%>

<% sub showHTMLHead() %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<meta name="GENERATOR" content="Hometown Code Generator 1.0">
<meta name="ProgId" content="<%=HTprogPrefix%>Add.asp">
<title>�s�W���</title>
<link rel="stylesheet" type="text/css" href="../inc/setstyle.css">
</head>
<body>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="100%" class="FormName" colspan="2"><%=HTProgCap%>&nbsp;<font size=2>�i<%=HTProgFunc%>�j</td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="FormLink" valign="top" colspan=2 align=right>
	       <a href="Javascript:window.history.back();">�^�e��</a>
		</td>
	  </tr>  
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
	  <tr>
	    <td align=center colspan=2 width=80% height=230 valign=top>    
<% end sub '--- showHTMLHead() ------%>


<% sub ShowHTMLTail() %>     
    </td>
  </tr>  
</table> 
</body>
</html>
<% end sub '--- showHTMLTail() ------%>


<% sub showErrBox() %>
	<script language=VBScript>
  			alert "<%=errMsg%>"
  			window.history.back
	</script>
<%
end sub '---- showErrBox() ----

sub checkDBValid()	'===================== Server Side Validation Put HERE =================

'---- ��ݸ�������ˬd�{���X��b�o��	�A�p�U�ҡA�����ɳ] errMsg="xxx" �� exit sub ------
'	SQL = "Select * From Client Where ClientID = '"& Request("pfx_ClientID") &"'"
'	set RSvalidate = conn.Execute(SQL)
'
'	if not RSvalidate.EOF Then
'		errMsg = "�u�Ȥ�s���v����!!�Э��s�إ߫Ȥ�s��!"
'		exit sub
'	end if

end sub '---- checkDBValid() ----

sub doUpdateDB()
