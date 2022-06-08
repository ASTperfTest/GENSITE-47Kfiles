<%@ CodePage = 65001 %>
<% Response.Expires = 0
HTProgCap="點閱統計"
HTProgFunc="點閱統計"
HTProgCode="GW1M22"
HTProgPrefix="mSession" %>
<% response.expires = 0 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbutil.inc" -->
<%
 Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>新網頁1</title>
</head>

<body>

<p>來源國別統計</p>
<form action="country_stat.asp" method="post">
<p>從<select size="1" name="y1">
  <option selected value="請選擇">請選擇</option>
  <option <% if request("y1")="2004" then response.write " selected" end if %>>2004</option>
  <option <% if request("y1")="2005" then response.write " selected" end if %>>2005</option>
  <option <% if request("y1")="2006" then response.write " selected" end if %>>2006</option>
  <option <% if request("y1")="2007" then response.write " selected" end if %>>2007</option>
  <option <% if request("y1")="2008" then response.write " selected" end if %>>2008</option>
  <option <% if request("y1")="2009" then response.write " selected" end if %>>2009</option>
  <option <% if request("y1")="2010" then response.write " selected" end if %>>2010</option>
</select>年<select size="1" name="m1">
  <option selected value="請選擇">請選擇</option>
  <option <% if request("m1")="1" then response.write " selected" end if %>>1</option>
  <option <% if request("m1")="2" then response.write " selected" end if %>>2</option>
  <option <% if request("m1")="3" then response.write " selected" end if %>>3</option>
  <option <% if request("m1")="4" then response.write " selected" end if %>>4</option>
  <option <% if request("m1")="5" then response.write " selected" end if %>>5</option>
  <option <% if request("m1")="6" then response.write " selected" end if %>>6</option>
  <option <% if request("m1")="7" then response.write " selected" end if %>>7</option>
  <option <% if request("m1")="8" then response.write " selected" end if %>>8</option>
  <option <% if request("m1")="9" then response.write " selected" end if %>>9</option>
  <option <% if request("m1")="10" then response.write " selected" end if %>>10</option>
  <option <% if request("m1")="11" then response.write " selected" end if %>>11</option>
  <option <% if request("m1")="12" then response.write " selected" end if %>>12</option>
</select>月<select size="1" name="d1">
  <option selected value="請選擇">請選擇</option>
  <option <% if request("d1")="1" then response.write " selected" end if %>>1</option>
  <option <% if request("d1")="2" then response.write " selected" end if %>>2</option>
  <option <% if request("d1")="3" then response.write " selected" end if %>>3</option>
  <option <% if request("d1")="4" then response.write " selected" end if %>>4</option>
  <option <% if request("d1")="5" then response.write " selected" end if %>>5</option>
  <option <% if request("d1")="6" then response.write " selected" end if %>>6</option>
  <option <% if request("d1")="7" then response.write " selected" end if %>>7</option>
  <option <% if request("d1")="8" then response.write " selected" end if %>>8</option>
  <option <% if request("d1")="9" then response.write " selected" end if %>>9</option>
  <option <% if request("d1")="10" then response.write " selected" end if %>>10</option>
  <option <% if request("d1")="11" then response.write " selected" end if %>>11</option>
  <option <% if request("d1")="12" then response.write " selected" end if %>>12</option>
  <option <% if request("d1")="13" then response.write " selected" end if %>>13</option>
  <option <% if request("d1")="14" then response.write " selected" end if %>>14</option>
  <option <% if request("d1")="15" then response.write " selected" end if %>>15</option>
  <option <% if request("d1")="16" then response.write " selected" end if %>>16</option>
  <option <% if request("d1")="17" then response.write " selected" end if %>>17</option>
  <option <% if request("d1")="18" then response.write " selected" end if %>>18</option>
  <option <% if request("d1")="19" then response.write " selected" end if %>>19</option>
  <option <% if request("d1")="20" then response.write " selected" end if %>>20</option>
  <option <% if request("d1")="21" then response.write " selected" end if %>>21</option>
  <option <% if request("d1")="22" then response.write " selected" end if %>>22</option>
  <option <% if request("d1")="23" then response.write " selected" end if %>>23</option>
  <option <% if request("d1")="24" then response.write " selected" end if %>>24</option>
  <option <% if request("d1")="25" then response.write " selected" end if %>>25</option>
  <option <% if request("d1")="26" then response.write " selected" end if %>>26</option>
  <option <% if request("d1")="27" then response.write " selected" end if %>>27</option>
  <option <% if request("d1")="28" then response.write " selected" end if %>>28</option>
  <option <% if request("d1")="29" then response.write " selected" end if %>>29</option>
  <option <% if request("d1")="30" then response.write " selected" end if %>>30</option>
  <option <% if request("d1")="31" then response.write " selected" end if %>>31</option>
</select>日至<select size="1" name="y2">
  <option selected value="請選擇">請選擇</option>
  <option <% if request("y2")="2004" then response.write " selected" end if %>>2004</option>
  <option <% if request("y2")="2005" then response.write " selected" end if %>>2005</option>
  <option <% if request("y2")="2006" then response.write " selected" end if %>>2006</option>
  <option <% if request("y2")="2007" then response.write " selected" end if %>>2007</option>
  <option <% if request("y2")="2008" then response.write " selected" end if %>>2008</option>
  <option <% if request("y2")="2009" then response.write " selected" end if %>>2009</option>
  <option <% if request("y2")="2010" then response.write " selected" end if %>>2010</option>
</select>年<select size="1" name="m2">
  <option selected value="請選擇">請選擇</option>
  <option <% if request("m2")="1" then response.write " selected" end if %>>1</option>
  <option <% if request("m2")="2" then response.write " selected" end if %>>2</option>
  <option <% if request("m2")="3" then response.write " selected" end if %>>3</option>
  <option <% if request("m2")="4" then response.write " selected" end if %>>4</option>
  <option <% if request("m2")="5" then response.write " selected" end if %>>5</option>
  <option <% if request("m2")="6" then response.write " selected" end if %>>6</option>
  <option <% if request("m2")="7" then response.write " selected" end if %>>7</option>
  <option <% if request("m2")="8" then response.write " selected" end if %>>8</option>
  <option <% if request("m2")="9" then response.write " selected" end if %>>9</option>
  <option <% if request("m2")="10" then response.write " selected" end if %>>10</option>
  <option <% if request("m2")="11" then response.write " selected" end if %>>11</option>
  <option <% if request("m2")="12" then response.write " selected" end if %>>12</option>
</select>月<select size="1" name="d2">
  <option selected value="請選擇">請選擇</option>
   <option <% if request("d2")="1" then response.write " selected" end if %>>1</option>
  <option <% if request("d2")="2" then response.write " selected" end if %>>2</option>
  <option <% if request("d2")="3" then response.write " selected" end if %>>3</option>
  <option <% if request("d2")="4" then response.write " selected" end if %>>4</option>
  <option <% if request("d2")="5" then response.write " selected" end if %>>5</option>
  <option <% if request("d2")="6" then response.write " selected" end if %>>6</option>
  <option <% if request("d2")="7" then response.write " selected" end if %>>7</option>
  <option <% if request("d2")="8" then response.write " selected" end if %>>8</option>
  <option <% if request("d2")="9" then response.write " selected" end if %>>9</option>
  <option <% if request("d2")="10" then response.write " selected" end if %>>10</option>
  <option <% if request("d2")="11" then response.write " selected" end if %>>11</option>
  <option <% if request("d2")="12" then response.write " selected" end if %>>12</option>
  <option <% if request("d2")="13" then response.write " selected" end if %>>13</option>
  <option <% if request("d2")="14" then response.write " selected" end if %>>14</option>
  <option <% if request("d2")="15" then response.write " selected" end if %>>15</option>
  <option <% if request("d2")="16" then response.write " selected" end if %>>16</option>
  <option <% if request("d2")="17" then response.write " selected" end if %>>17</option>
  <option <% if request("d2")="18" then response.write " selected" end if %>>18</option>
  <option <% if request("d2")="19" then response.write " selected" end if %>>19</option>
  <option <% if request("d2")="20" then response.write " selected" end if %>>20</option>
  <option <% if request("d2")="21" then response.write " selected" end if %>>21</option>
  <option <% if request("d2")="22" then response.write " selected" end if %>>22</option>
  <option <% if request("d2")="23" then response.write " selected" end if %>>23</option>
  <option <% if request("d2")="24" then response.write " selected" end if %>>24</option>
  <option <% if request("d2")="25" then response.write " selected" end if %>>25</option>
  <option <% if request("d2")="26" then response.write " selected" end if %>>26</option>
  <option <% if request("d2")="27" then response.write " selected" end if %>>27</option>
  <option <% if request("d2")="28" then response.write " selected" end if %>>28</option>
  <option <% if request("d2")="29" then response.write " selected" end if %>>29</option>
  <option <% if request("d2")="30" then response.write " selected" end if %>>30</option>
  <option <% if request("d2")="31" then response.write " selected" end if %>>31</option>
</select>日&nbsp; <input type="submit" value="開始統計" name="B3"></p>
</form>
<%
  date1=request("y1") & "/" & request("m1") & "/" & request("d1")
  date2=request("y2") & "/" & request("m2") & "/" & request("d2")
  if cint(isdate(date1))<>0 and cint(isdate(date2))<>0 then
  sql="SELECT COUNT(*) AS Expr1,IPCat FROM gipHitSession WHERE (hitFirst BETWEEN N'" & date1 & "' AND N'" & date2 & "') GROUP BY  IPCat"
  set rs=conn.Execute(sql)
%>
<table border="1" width="100%">
  <tr>
    <td width="23%">國別來源</td>
    <td width="77%">訪問次數</td>
  </tr>
  <% do while not rs.eof  %> 
  <tr>
    <td width="23%"><% =rs(1) %>&nbsp;</td>
    <td width="77%"><% =rs(0) %>&nbsp;</td>
  </tr>
<%  rs.MoveNext
    loop
%>    
</table>
<p><input type="button" value="轉成Excel" name="B3"></p>
<%
  end if
%>
</body>

</html>
