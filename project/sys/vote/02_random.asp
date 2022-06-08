<%@ CodePage = 65001 %>
<%
	Response.Expires = 0
	'HTProgCap = "功能管理"
	'HTProgFunc = "問卷調查"
	HTProgCode = "GW1_vote01"
	'HTProgPrefix = "mSession"
	response.expires = 0
%>
<!-- #include virtual = "/inc/client.inc" -->
<!-- #include virtual = "/inc/dbutil.inc" -->
<%
	Set RSreg = Server.CreateObject("ADODB.RecordSet")
%>
<%	

	subjectid = request("subjectid")
	prizeno = request("prizeno" & subjectid)

	
	if subjectid = "" then
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	
	sql = "" & _
		" select isNull(m011_haveprize, '0'), isNull(m011_pflag, '0') " & _
		" from m011 where m011_subjectid = " & subjectid
	set rs = conn.execute(sql)
	if rs.EOF then
		response.write "<script language='JavaScript'>history.go(-1);</script>"
		response.end
	end if
	haveprize = rs(0)
	pflag = rs(1)
	
	if haveprize = "0" then
		response.write "<script language='JavaScript'>alert('此題不提供抽獎！');history.go(-1);</script>"
		response.end
	end if
	
	set rs2 = conn.execute("select m014_id from m014 where m014_subjectid = " & subjectid)
	if rs2.EOF then
		response.write "<script language='JavaScript'>alert('無任何答題者資料！');history.go(-1);</script>"
		response.end
	end if	
	
	if pflag = "0" then		' pflag = 0, 表示此題尚未抽獎

		m014_id_str = ","
		while not rs2.EOF
			m014_id_str = m014_id_str & rs2(0) & ","
			rs2.MoveNext
		wend
		
		if m014_id_str <> "" then
			m014_id_str = mid(m014_id_str, 1, len(m014_id_str) - 1)
			m014_id = split(m014_id_str, ",")
			
			array_ubound = UBound(m014_id)
			
			for i = 1 to array_ubound
				Randomize
				random_no = (CInt(Rnd * 1000 * second(time)) mod array_ubound) + 1
				
				tmp = m014_id(i)
				m014_id(i) = m014_id(random_no)
				m014_id(random_no) = tmp
			next
			
			if CInt(prizeno) > CInt(array_ubound) then
				prizeno = array_ubound
			end if
			
			update_m014_id = ""
			for i = 1 to prizeno		' 取得亂數後被選到的 m014_id
				update_m014_id = update_m014_id & m014_id(i) & ","
			next
			update_m014_id = mid(update_m014_id, 1, len(update_m014_id) - 1)
			
			
			sql = "" & _
				" update m014 set m014_pflag = '0' where m014_subjectid = " & subjectid & "; " & _
				" update m014 set m014_pflag = '1' where m014_subjectid = " & subjectid & _
				" and m014_id in (" & update_m014_id & "); " & _
				" update m011 set m011_pflag = '1' where m011_subjectid = " & subjectid
			conn.execute(sql)
			
			response.write "<script language='JavaScript'>alert('已抽出 " & prizeno & " 位中獎人！');location.replace('02_random.asp?subjectid=" & subjectid & "&functype=" & functype & "');</script>"
			response.end
		end if
		
	end if
	
	

	set rs = conn.execute("select m011_subject from m011 where m011_subjectid = " & subjectid)
	subject = trim(rs(0))
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="style.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<p>問卷調查維護－亂數抽取問卷答卷得獎者基本資料</p>
<table width="100%" border="1" cellspacing="1" cellpadding="5">
  <tr> 
    <td>問卷題目：</td>
    <td><p><%= subject %></p></td>
  </tr>
</table>
<p>得獎者資料：</p>
<table width="100%" border="1" cellspacing="1" cellpadding="5">
<%
	sql = "select m014_name, m014_email from m014 where m014_subjectid = " & subjectid & _
		" and m014_pflag = '1' "
	set rs = conn.execute(sql) 
	while not rs.EOF
%>
  <tr> 
    <td>姓名：</td>
    <td><%= trim(rs(0)) %></td>
  </tr>
  <tr> 
    <td>email：</td>
    <td><a href="mailto:<%= trim(rs(1)) %>"><%= trim(rs(1)) %></a></td>
  </tr>
<%
		rs.MoveNext
	wend
%>
</table>
<p><a href="javascript:location.href='02.asp';">Previous Page</a></p>
</body>
</html>
