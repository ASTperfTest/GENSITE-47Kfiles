<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="關鍵字詞維護"
HTProgFunc="編修關鍵字詞"
HTProgCode="GC1AP7"
HTProgPrefix="KeywordWTP" %>
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
if request("task")="編修存檔" then
	xDate=cStr(date())
	xKeyword=pkStr(request("xKeyword"),"")
	xKeywordNew=pkStr(request("xKeywordNew"),"")	
	SQLAct="Insert Into CuDTKeywordWTP values("+xKeyword+",0,"+xKeywordNew+",'E','"+session("UserID")+"','"+xDate+"')"
	Set RS = Conn.execute(SQLAct)
%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	</head>
		<script language=VBScript>
		  alert("已列入關鍵字詞待處理清單中！")
		  window.navigate "<%=HTProgPrefix%>List.asp?nowPage=1"
		</script>	
	</html>
<%     		response.end
elseif request("task")="刪除" then
	xDate=cStr(date())
	xKeyword=pkStr(request("xKeyword"),"")	
	SQLAct="Insert Into CuDTKeywordWTP values("+xKeyword+",0,null,'D','"+session("UserID")+"','"+xDate+"')"
	Set RS = Conn.execute(SQLAct)
%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	</head>
		<script language=VBScript>
		  alert("已列入關鍵字詞待處理清單中！！")
		  window.navigate "<%=HTProgPrefix%>List.asp?nowPage=1"
		</script>	
	</html>
<%     		response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8; no-caches">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>關鍵字詞維護</title>
</head>
<body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<table border="0" width="100%" cellspacing="1" cellpadding="0">
  <tr>
	    <td width="50%" class="FormName" align="left"><%=HTProgCap%>&nbsp;<font size=2>【<%=HTProgFunc%>】</td>
		<td width="50%" class="FormLink" align="right">
			<a href="Javascript:window.history.back()">回前頁</a>	
	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
  <tr>
    <td class="Formtext" colspan="2" height="15"></td>
  </tr>
  <tr>
    <td width="95%%" colspan="2">
      <form method="POST" name="reg">
       <input type=hidden name="task">              
          <center>    
          <table border="0" width="90%" cellspacing="1" cellpadding="3" class="bluetable">    
            <tr>    
              <td class="lightbluetable" align="right">原關鍵字詞</td>    
              <td class="whitetablebg"><input type="text" name="xKeyword" size="50" maxlength=50></td>    
    	    </tr>                                                                                
            <tr>    
              <td class="lightbluetable" align="right">新關鍵字詞</td>    
              <td class="whitetablebg"><input type="text" name="xKeywordNew" size="50" maxlength=50></td>    
    	    </tr> 
          </table>
          </center>
            <p align="center">  
            <img src="../images/management/titlehr.gif" width="570" height="1" vspace="3"><br>
            <% if (HTProgRight and 8)=8 then %><input type="button" value="編修存檔" class="cbutton" onclick="VBS: formEdit"><% end if %>&nbsp;
            <% if (HTProgRight and 16)=16 then %><input type="button" value="刪除" class="cbutton" onclick="VBS: formDelete"><% end if %>
            					<input type="reset" value="清除重填" class="cbutton">         
      </form>          
    </td>
  </tr>  
</table>
</body>
</html> 
<script language="vbscript">
Sub formEdit
    if reg.xKeyword.value="" then
    	alert "請輸入原關鍵字詞!"
    	reg.xKeyword.focus
    	exit sub
    end if
 
	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_keyword.asp?xKeyword=" & B5toUTF8(reg.xKeyword.value)
	oXML.load(xURI)	
	set xKeywordNode = oXML.selectSingleNode("xKeywordList/xKeywordStr")
	if xKeywordNode.text<>"" then      
    		alert "原關鍵字詞不存在, 無法編修!"
    		exit sub
    	end if  
    if reg.xKeywordNew.value="" then
    	alert "請輸入新關鍵字詞!"
    	reg.xKeywordNew.focus
    	exit sub
    end if    	
    if blen(reg.xKeywordNew.value)<=2 then
    	alert "新關鍵字詞至少二個字以上!"
    	reg.xKeywordNew.focus
    	exit sub
    end if   
	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_keyword.asp?xKeyword=" & B5toUTF8(reg.xKeywordNew.value)
	oXML.load(xURI)	
	set xKeywordNode = oXML.selectSingleNode("xKeywordList/xKeywordStr")    
	if xKeywordNode.text<>"" then      
    		reg.Task.value = "編修存檔"    
    		reg.Submit  
    	else
    		alert "新關鍵字詞已存在!"
    		exit sub
    	end if  	 	
End Sub
Sub formDelete
    if reg.xKeyword.value="" then
    	alert "請輸入原關鍵字詞!"
    	reg.xKeyword.focus
    	exit sub
    end if

	set oXML = createObject("Microsoft.XMLDOM")
	oXML.async = false
	xURI = "<%=session("mySiteMMOURL")%>/ws/ws_keyword.asp?xKeyword=" & B5toUTF8(reg.xKeyword.value)
	oXML.load(xURI)	
	set xKeywordNode = oXML.selectSingleNode("xKeywordList/xKeywordStr")
	if xKeywordNode.text<>"" then      
    		alert "原關鍵字詞不存在, 無法刪除!"
    		exit sub
    	else
            chky=msgbox("注意！"& vbcrlf & vbcrlf &"　確定刪除此筆資料嗎？"& vbcrlf , 48+1, "請注意！！")
            if chky=vbok then
    		reg.Task.value = "刪除"    
    		reg.Submit  
    	    end if
    	end if  
End Sub
</script> 
<script language=javascript>
function B5toUTF8(x){
	return encodeURI(x);
}
</script>