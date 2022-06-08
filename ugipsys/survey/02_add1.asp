<%@ CodePage = 65001 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%

	subjectid = request("subjectid")
	if subjectid = "" then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	
	sql = "update m011 set m011_subjectid=" & pkstr(subjectid,"") _
		& " where giCuItem=" & pkstr(subjectid,"")
	conn.execute sql
	sql = "select m011_questno, m011_jumpquestion from m011 where giCuItem = " & pkstr(subjectid,"")
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	questno = rs(0)
	jumpquestion = rs(1)
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" href="css.css">
</head>

<body bgcolor="#FFFFFF">
  <tr>
	    <td width="100%" class="FormName" colspan="5">單元資料維護&nbsp;
	    <font size=2>【主題單元: 問卷調查】</td>
  </tr>
  <tr>
    <td width="100%" colspan="5">
      <hr noshade size="1" color="#000080">
    </td>
  </tr>
  <tr>
    <!--td class="FormLink" valign="top" align=right><a href="http:adm_inves.asp" title="廣告列表">問卷列表</a--> <!--a href="adm_inves.asp" title="新增廣告">檢視投票結果</a--></td>
  </tr>
  <tr>
    <td class="FormLink" valign="top"><strong><!--font color="#990000">新增問卷</font></strong--><p>（步驟二：新增問卷主題）</p> </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>
  </tr>
<table border="1" cellspacing="0" cellpadding="4" width="100%" bordercolordark="#FFFFFF">
  <tr> 
      <td colspan="3">
      <p><font color="#993300">使用說明：</font></p>
      <ul><li><font color="#993300">請先輸入每一題目的答案數。</font></li></ul>
      </td>
  </tr>
  <tr> 
    <td colspan="3" align="center" valign="top" > 
      <table border="1" cellspacing="0" cellpadding="2" width="100%" bordercolordark="#FFFFFF">

      <form method="post" name="form1" action="02_add2_jump.asp">
      <input type="hidden" name="subjectid" value="<% =subjectid %>">
      <input type="hidden" name="questno" value="<% =questno %>">
      <input type="hidden" name="jumpquestion" value="<% =jumpquestion %>">
        
<%
	for i = 1 to questno
		sql = "select m012_answerno from m012 where m012_subjectid = " & subjectid & _
			" and m012_questionid = " & i
		set rs = conn.execute(sql)
		if rs.EOF then
			answerno = 0
		else
			answerno = rs(0)
		end if
%>
        <tr> 
          <td>題目<% =i %>：</td>
          <td>答案數：
            <select name="answerno<% =i %>">
<%
		for j = 1 to 10
			if j = answerno then
				response.write "<option value='" & j & "' selected>" & j & "</option>"
			else
				response.write "<option value='" & j & "'>" & j & "</option>"
			end if
		next
%>
            </select>
          </td>
        </tr>
<%
	next
%>
        <tr> 
          <td colspan="2">
            <input type="button" name="back" value="回上一步" onClick="window.location='adm_inves_new_fix.asp?subjectid=<% =subjectid %>'">
            <input type="submit" name="submit" value="下一步(問卷設計)">
          </td>
        </tr>
      </form>
      </table>
      <p>&nbsp;</p>
    </td>
  
  </tr>
 
</table>
</body>
</html>
