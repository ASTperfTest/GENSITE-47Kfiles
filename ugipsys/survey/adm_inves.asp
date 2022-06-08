<%@ CodePage = 65001 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<%
  set ts = conn.Execute("select count(*) from m011")
	
	totalrecord = ts(0)
	if totalrecord > 0 then              'ts代表筆數
		totalpage = totalrecord \ 10
		if (totalrecord mod 10) <> 0 then
			totalpage = totalpage + 1
		end if
	else
		totalpage = 1
	end if
	
	if request("page") = empty then
		page = 1
	else
		page = request("page")
	end if
'=================================================================================================
%>
<Html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8;no-caches;">
<link rel="stylesheet" href="inc/setstyle.css">
<title>問卷調查</title>
</head>
<body>
<table border="0" width="100%" cellspacing="1" cellpadding="0" align="center">
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
    <td class="FormLink" valign="top" align=right><a href="http://hywade.hyweb.com.tw/GipEdit/ctXMLin.asp?ItemID=6&CtNodeID=61" title="問卷列表">問卷列表</a> <!--a href="adm_inves_new.asp" title="新增廣告">新增問卷</a--></td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>
  </tr>
  <tr>
    <td class="FormLink" valign="top"><form action="" method="get">
    <table width="100%" border="0" cellpadding="5" cellspacing="1">
      <tr class="eTableLable">
        <td width="3%" height="31" align="center" nowrap><strong>編<br>
      號</strong></td>
        
        <td width="16%" align="center" nowrap><strong>問卷調查起訖日期</strong></td>
        <td width="9%" align="center" nowrap><strong>答題</strong></td>
        <td width="44%" align="center" nowrap><strong>問卷主題</strong></td>
        <td width="5%" align="center" nowrap><strong>結果</strong></td>
        <td width="13%" align="center" nowrap><strong>亂數抽獎</strong></td>
      </tr>
   <%
	sql = "" & _
		" select m011_subjectid, m011_subject, m011_bdate, m011_edate, m011_jumpquestion, " & _
		" m011_haveprize, isNull(m011_pflag, '0') m011_pflag " & _
		" from m011 order by m011_bdate desc "
	set rs = conn.execute(sql)
	i = 1
	do while not rs.EOF
		if i <= (page*10) and i > (page-1)*10 then
			if rs("m011_jumpquestion") = "1" then
				jumpstr = "跳題"
			else
				jumpstr = "一般"
			end if
%>      
      <tr class="eTableContent">
        <td align="center" bgcolor="#FFFFFF"><% =i %></td>
        <td align="center" bgcolor="#FFFFFF"><% =rs("m011_bdate") %> <br>
      ~<br>
      <% =rs("m011_edate") %></td>
        <td align="center" bgcolor="#FFFFFF"><% =jumpstr %></td>
        <td><a href="adm_inves_new_fix.asp?subjectid=<% =rs("m011_subjectid") %>"><% =trim(rs("m011_subject")) %></a></td>
        <td align="center"><a href="02_result.asp?subjectid=<% =rs("m011_subjectid") %>">結果</a></td>
<%
			if rs("m011_haveprize") = "1" then
				if rs("m011_pflag") = "0" then
%>
        <td align="center" nowrap bgcolor="#FFFFFF">
          <select name="prizeno<%= rs("m011_subjectid") %>">
<%
					for j = 1 to 5
						response.write "<option value='" & j & "'>" & j & "</option>"
					next
%>
          </select>
          <input name="Submit" type="submit" class="button" value="抽獎">
        </td>
<%
				elseif rs("m011_pflag") = "1" then
%>
          <td>
            <a href="02_random.asp?subjectid=<% =rs("m011_subjectid") %>">得獎名單</a>
          </td>
<%				
				end if
			else
				response.write "<td>此題不抽獎</td>"
			end if
%>
        </tr>
<%
		end if
		i = i + 1
		rs.MoveNext
	loop
%>

    </table>
    </form></td>
  </tr>
  <tr>
    <td class="FormLink" valign="top">&nbsp;</td>    
    </tr>
</table> 

</body>
</html>                                 


<script Language=VBScript>
  dim gpKey
     sub GoPage_OnChange      
           newPage=reg.GoPage.value     
           document.location.href="DsdXMLList.asp?nowPage=" & newPage & "&strSql=" & strSql & "&pagesize=15"    
     end sub      
     
     sub PerPage_OnChange                
           newPerPage=reg.PerPage.value
           document.location.href="DsdXMLList.asp?nowPage=1" & "&strSql=SELECT+htx%2E%2A%2C+ghtx%2E%2A%2C+xref1%2EmValue+AS+xref1fCTUPublic%2C+xref2%2EmValue+AS+xref2topCat+FROM+%28%28CuDTx7+AS+htx++JOIN+CuDtGeneric+AS+ghtx+ON+ghtx%2EiCUItem%3Dhtx%2EgiCUItem++LEFT+JOIN+codeMain+AS+xref1+ON+xref1%2EmCode+%3D+ghtx%2EfCTUPublic+AND+xref1%2EcodeMetaID%3D%27isPublic%27%29+LEFT+JOIN+codeMain+AS+xref2+ON+xref2%2EmCode+%3D+ghtx%2EtopCat+AND+xref2%2EcodeMetaID%3D%27topDataCat%27%29+WHERE+2%3D2++AND+ghtx%2EiCtUnit%3D20++ORDER+BY+xPostDate+DESC&pagesize=" & newPerPage                    
     end sub 

     sub setpKey(xv)
     	gpKey = xv
     end sub

</script>
