<%@ CodePage = 65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
today_str = date
	today_y = Year(today_str)
	today_m = Month(today_str)
	today_d = Day(today_str)

subjectid = request("subjectid")

	if subjectid = "" then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	
	
	sql = "update m011 set m011_edate=g.xPostDateEnd, m011_subject=g.sTitle, m011_subjectid=g.iCuItem" _
		& " FROM m011 as m JOIN CuDtGeneric as g on m.giCuItem=g.iCuItem where giCuItem = " & subjectid
	conn.execute sql
	
	sql = "select giCuItem, m011_subject, m011_questno, m011_bdate, m011_edate, m011_notetype, " & _
		" m011_haveprize, m011_jumpquestion, m011_onlyonce" & _
		" from m011 where giCuItem = " & subjectid
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<body onload='javascript:history.go(-1);'>"
		response.end
	end if
	notetype = trim(rs("m011_notetype"))
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="inc/setstyle.css">
<title>問卷調查</title>
</head>
<script language="JavaScript" type="text/JavaScript">
<!--
function CheckValue() {
	if ( document.form1.subject.value == '' ) {
		alert('請輸入主題名稱！');
		document.form1.subject.focus();
		return false;
	}
	
	return true;
}
//-->
</script>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
<form method="post" name="form1" action="02_add_act.asp" onSubmit="return CheckValue();">
<input type="hidden" name="subjectid" value="<% =subjectid %>">
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
    <td class="FormLink" valign="top"><strong><!--font color="#990000">新增問卷</font></strong--><p>（步驟一：建立問卷基本資料）</p> </td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">
      <table width="100%" border="0" cellpadding="5" cellspacing="1">
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>問卷主題：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <input name="subject" type="text" class="inputbox" size="60" value="<% =replace(trim(rs("m011_subject")), chr(34), "&quot;") %>">
          </td>
        </tr>
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>題目數量：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <select name="questno">
<%
	for i = 1 to 20
		if CInt(i) = CInt(rs("m011_questno")) then
			response.write "<option value='" & i & "' selected>" & i & "</option>"
		else
			response.write "<option value='" & i & "'>" & i & "</option>"
		end if
	next
%>

            </select> 
      題</td>
        </tr>
        <!--tr>
          <td align="right" nowrap class="eTableLable"><strong>答題方式：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <input name="jumpquestion" type="radio" value="0" <% if rs("m011_jumpquestion") = "0" then %> checked<% end if %>>
      一般（ 逐題作答）　
      
<input type="radio" name="jumpquestion" value="1" <% if rs("m011_jumpquestion") = "1" then %> checked<% end if %>>
      跳題</td>
        </tr-->
        <tr>
          <td align="right" valign="top" nowrap class="eTableLable"><strong>填卷者資料：</strong></td>
                <td valign="top" nowrap>
                  <label>
                  <input type="radio" name="haveprize" value="0" <% if rs("m011_haveprize") = "0" then %> checked<% end if %>>
            不需要</label>
                  <label>
                  <input name="haveprize" type="radio" value="1" <% if rs("m011_haveprize") = "1" then %> checked<% end if %>>
            需要 （請勾選所需資料）</label>
            <table border="0" cellpadding="5" cellspacing="0" class="eTableContent">
              <tr>
                <td>
                    <table width="100%" border="0" cellpadding="3" cellspacing="3" class="eTableContent">
                      <tr>
                        <td><input type="checkbox" name="notetype" value="1" <% if mid(notetype, 1, 1) = "1" then %> checked<% end if %>>姓名</td>
                        <td><input type="checkbox" name="notetype" value="2" <% if mid(notetype, 2, 1) = "1" then %> checked<% end if %>>聯絡方式</td>
                        <td><input type="checkbox" name="notetype" value="3" <% if mid(notetype, 3, 1) = "1" then %> checked<% end if %>>電話</td>
                        <td><input type="checkbox" name="notetype" value="4" <% if mid(notetype, 4, 1) = "1" then %> checked<% end if %>>E-mail</td>
                      </tr>
                      <!--tr>
                        <td><input type="checkbox" name="notetype" value="5" <% if mid(notetype, 5, 1) = "1" then %> checked<% end if %>>居住縣市</td>
                        <td><input type="checkbox" name="notetype" value="6" <% if mid(notetype, 6, 1) = "1" then %> checked<% end if %>>家庭成員數</td>
                        <td><input type="checkbox" name="notetype" value="7" <% if mid(notetype, 7, 1) = "1" then %> checked<% end if %>>收入</td>
                        <td><input type="checkbox" name="notetype" value="8" <% if mid(notetype, 8, 1) = "1" then %> checked<% end if %>>職業</td-->
                      </tr-->
                    </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td align="right" nowrap class="eTableLable"><strong>填卷限制：</strong></td>
          <td bgcolor="#FFFFFF" class="eTableContent">
            <label>
            <input name="onlyonce" type="radio" value="0" <% if rs("m011_onlyonce") = "0" then %> checked<% end if %>>
    不限制
    <input type="radio" name="onlyonce" value="1" <% if rs("m011_onlyonce") = "1" then %> checked<% end if %>>
    每人限填一次</label>
            <label></label>
          </td>
        </tr>

      </table>
      <table width="100%" border="0" align="center" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <tr>
          <td width="60%" height="40" align="right"> <!--a href="javascript:history.go(-1);">back</a-->
	      <input type="button" name="back" value="回上一步" onClick="javascript:history.go(-1);">          
              <input name="submit" type="submit"  value="繼續修改問卷內容">
              <!--input type="submit" name="submit" value="刪除本問卷" onClick="return confirm('您確定刪除？');"-->
          </td>
        </tr>
      </table>
</td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>    
    </tr>
</form>
</table> 

</body>
</html>                                 


    
     
  
     