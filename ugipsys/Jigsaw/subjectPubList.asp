<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="主題館績效統計"
HTProgFunc="查詢清單"
HTProgCode="jigsaw"
HTProgPrefix="newkpi" %>

<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/checkQS.inc" -->
 <%

 call CheckURL(Request.QueryString)
session("jigsql") = ""
session("jigcheck")=""
session("jigcheck1")=""
'response.write request("iCUItem")

sql="SELECT stitle FROM [mGIPcoanew].[dbo].[CuDTGeneric] where iCUItem='"&request("iCUItem")&"'"
Set rs = conn.Execute(sql)
sql1="SELECT KnowledgeJigsaw.gicuitem as gicuitem, CuDTGeneric.sTitle as sTitle, CuDTGeneric.fCTUPublic as fCTUPublic, KnowledgeJigsaw.orderSiteUnit as orderSiteUnit, KnowledgeJigsaw.orderSubject as orderSubject, KnowledgeJigsaw.orderKnowledgeTank as orderKnowledgeTank, KnowledgeJigsaw.orderKnowledgeHome as orderKnowledgeHome, KnowledgeJigsaw.parentIcuitem as parentIcuitem , KnowledgeJigsaw.gicuitem as gicuitem FROM  KnowledgeJigsaw INNER JOIN CuDTGeneric ON KnowledgeJigsaw.gicuitem = CuDTGeneric.iCUItem where KnowledgeJigsaw.parentIcuitem='"&request("iCUItem")&"'"
Set rs1 = conn.Execute(sql1)

'response.write "sql=" & sql & "<br/>"
'response.write "sql1=" & sql1

%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
	<link rel="stylesheet" href="/inc/setstyle.css">
	<link type="text/css" rel="stylesheet" href="../css/list.css">
	<link type="text/css" rel="stylesheet" href="../css/layout.css">
	<link type="text/css" rel="stylesheet" href="../css/setstyle.css">
<title></title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
  <tr>
	    <td width="50%" align="left" nowrap class="FormName">農業推薦單元知識拼圖區塊管理&nbsp;
		<font size=2>【內容區塊--<%=rs(0)%>主題專區】</font>
		</td>
		<td width="50%" class="FormLink" align="right">
			<!--A href="subjectPubAdd(2).htm" title="新增">新增區塊</A-->
			<A href="Javascript:window.history.back();" title="回前頁">回前頁</A>			
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
    <td width="95%" colspan="2" height=230 valign=top>
 <Form name=reg method="POST" action="subjectPubList01.asp?iCUItem=<%=request("iCUItem")%>">
  
    <CENTER>
	  <table width=100% cellpadding="0" cellspacing="1" id="ListTable">
        <tr align=left>
          <th>專區區塊條列標題</th>
          <th width="15%">是否公開</th>
          <th width="25%">內容文章設定</th>
        </tr>
        <%
      Do Until rs1.eof
        if rs1("sTitle")<>"議題關聯知識文章單元順序設定" then
		%>
		<tr>
          <td class=eTableContent><%=rs1("sTitle")%></td>
          <td align="right" class=eTableContent>
		  
		  <select  name="<%=rs1("gicuitem")%>" id="<%=rs1("gicuitem")%>" >
           
				<option value="Y"  <%if rs1("fCTUPublic")="Y" then response.write "selected" end if%>>公開</option>
        <option value="N"  <%if rs1("fCTUPublic")="N" then response.write "selected" end if%>>不公開</option>
      </select></td>
            <td class=eTableContent align="right">		    
			    <%select case rs1("sTitle") %>
			    <%case "最新議題" %>
			        <input name="button" type="button" class="cbutton" onClick="location.href='subject_setList(2).asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=rs1("gicuitem")%>'" value="內容管理">
			        <input name="button6" type="button" class="cbutton" onClick="location.href='subject_query.asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=rs1("gicuitem")%>&latest=1'" value="新增內容">
			    <%case "資源推薦的超連結" %>    
			        <input name="button" type="button" class="cbutton" onClick="location.href='subject_setList(2).asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=rs1("gicuitem")%>&ActionType=setLink'" value="內容管理">					
				<%case "使用者參與討論或分享心得" %>
				<input name="button" type="button" class="cbutton" onClick="location.href='subject_setList(2).asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=rs1("gicuitem")%>&ActionType=ManageDiscussion'" value="留言管理">
				
			    <%case else %>
			        <input name="button" type="button" class="cbutton" onClick="location.href='subject_setList(2).asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=rs1("gicuitem")%>'" value="內容管理">
			        <input name="button6" type="button" class="cbutton" onClick="location.href='subject_query.asp?iCUItem=<%=request("iCUItem")%>&gicuitem=<%=rs1("gicuitem")%>'" value="新增內容">
			    <%end select %>
		    </td>
        </tr>
        <%else%>
        <tr>
          <td colspan="3" class=eTableContent>
		  <%=rs1("sTitle")%>：
		  入口網(站內單元) <input name="orderSiteUnit"value="<%=rs1("orderSiteUnit")%>" size="5">
		  主題館 <input name="orderSubject" value="<%=rs1("orderSubject")%>" size="5">
		  知識庫 <input name="orderKnowledgeTank" value="<%=rs1("orderKnowledgeTank")%>" size="5">
		  知識家 <input name="orderKnowledgeHome" value="<%=rs1("orderKnowledgeHome")%>" size="5">
		  </td>
        </tr>
       <%
	   end if
	   rs1.MoveNext
		Loop
	%>
     
      </table>
    </CENTER>
      
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
 <td align="center">     
        <input name="button4" type="submit" class="cbutton" onClick="location.href='subjectPubList01.asp?<%=request("iCUItem")%>'" value="編修存檔">
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="location.href='index.asp'">
 </td>
</tr>
</table>
     </form>  
    </td>
  </tr>  
	<tr>
		<td width="100%" colspan="2" align="center">
	</td></tr>
</table> 

</body>
</html>                                 
<script language="JavaScript">
function ret(sTitle,fCTUPublic){

window.opener.document.OAConfigForm.allowUserid.value=sTitle;
window.opener.document.OAConfigForm.allowUserName.value=fCTUPublic;
window.opener.document.OAConfigForm.allowUserName.value=orderSiteUnit;
window.opener.document.OAConfigForm.allowUserName.value=orderSubject;
window.opener.document.OAConfigForm.allowUserName.value=orderKnowledgeTank;
window.opener.document.OAConfigForm.allowUserName.value=fCTUPublic;
//window.opener.document.getElementById("username").value=username;
window.close();
}
</script >
