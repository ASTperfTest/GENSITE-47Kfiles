<%@ CodePage = 65001 %><!-- saved from url=(0022)http://internet.e-mail -->
<html>
<head>
<title>HTSDModel</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="setstyle.css">
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0" marginheight="0" marginwidth="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr background="images/namebg.gif"> 
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" background="images/namebg1.gif">
        <tr>
          <td width="89%">
          <DIV style="filter:Shadow(Direction=135,color='#336633');height:35">
          <b><font face="新細明體" color=#eeeeee size=6><%=session("mySiteName")%> 後台管理系統</font></b>
          <td align="right" width="11%"><img src="images/name5.gif" width="116" height="59"></td>
        </tr>
      </table> 
    </td>
  </tr>
  <tr> 
    <td bgcolor="#7A9266"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="20"> 
            <table border="0" cellspacing="0" cellpadding="0" bgcolor="#7A9266">
	        <tr>
	          <td width="20" align="center"><img ID=menuimg src="images/X-2.gif" alt="收放視窗" width="13" height="13" style="cursor: hand;" onclick="VBScript: menuimgon"></td>
		</tr> 
	    </table>
	  </td>	               
            <td colspan="2" valign="baseline">&nbsp; </td>
          <td align="right" valign="baseline" width="10%">
          <img src="images/home.gif" width="19" height="20" alt="回首頁" style="cursor: hand; " onclick="VBS:window.top.mainFrame.location='../main.htm'">                    
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#4B5C3E"><img src="images/shim.gif" width="5" height="5"></td>
  </tr>
</table>
</body>
</html>
<script language=VBS>
Sub menuimgon()
 	if window.parent.f.cols="0,*" Then
		window.parent.f.cols="159,*"
		menuimg.Src="images/x-2.gif"
	Else
		window.parent.f.cols="0,*"
		menuimg.Src="images/x-1.gif"
	End If
 End Sub
sub CalendarClick
	window.parent.frames(2).navigate "Calendar.asp"
end sub
</script>
