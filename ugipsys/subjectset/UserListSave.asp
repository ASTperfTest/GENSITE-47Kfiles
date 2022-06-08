<%@ CodePage = 65001 %>
<% Response.Expires = 0
	response.charset="utf-8"
   HTProgCode = "webgeb1" %>
<!--#include virtual = "/inc/index.inc" -->
<!--#include virtual = "/inc/server.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->

<%
if request("ckbox")= "" and request("type") = 1 then
	Response.Write "<script language='javascript'>"
	Response.Write "alert('請選擇至少一名專家帳號！');"
	Response.Write "history.go(-1);"
	Response.Write "</script>"
end if
'response.write request("ckbox") &"<br>" '使用者帳號
'response.write request("id")&"<br>" '主題館id
'response.write request("type")&"<br>"	
'response.write request("deptid")&"<br>"	
'response.end
'原來的使用者
sql_ch = "select NodeInfo.subject_user,NodeInfo.owner,InfoUser.deptID from "
sql_ch = sql_ch & " NodeInfo LEFT OUTER JOIN InfoUser ON NodeInfo.owner = InfoUser.UserID "
sql_ch = sql_ch & " where NodeInfo.CtrootID='"&request("id")&"'"
set rs_ch = conn.execute(sql_ch)
o_subject_user= rs_ch("subject_user")
o_owner= rs_ch("owner")
o_owner_dept = rs_ch("deptID")
'管理權限
if request("type") = 1 then
	new_owner = request("ckbox")
	'更新管理權限使用者
	sql_update = "UPDATE NodeInfo SET owner = '"& new_owner &"' WHERE CtrootID = '"& request("id") &"'"
	conn.execute(sql_update)
	'移除原管理者的上稿權限
	if  not (instr(o_subject_user,o_owner) > 0) then
		'取主題館所有的nodeid
		sql2 = "SELECT CtNodeID FROM CatTreeNode WHERE CtRootID = '"&request("id")&"'"
		set rs3 = conn.execute(sql2)
		while not rs3.eof
			sql_del = "UPDATE CtUserSet SET Rights = '0' where UserID = '"& o_owner &"' and CtNodeID = '"&rs3("CtNodeID")&"'"
			conn.execute(sql_del)
			rs3.movenext
		wend
	end if
	'轉移文章
	'取主題館所有的CtUnitID
	sql = "SELECT CtUnitID FROM CatTreeNode WHERE CtRootID = '"&request("id")&"'"
	set rs_ctunit = conn.execute(sql)
	while not rs_ctunit.eof
		if rs_ctunit("CtUnitID") <> "" then
			sql_update2 = "UPDATE CuDTGeneric SET iEditor = '"& new_owner &"' ,iDept ='"& request("deptid")&"' WHERE iCTUnit = '"& rs_ctunit("CtUnitID") &"' and  iEditor ='"& o_owner &"'"
			conn.execute(sql_update2)
			'response.write sql_update2&"<br>"
		end if
		rs_ctunit.movenext
	wend
	'增加上稿權限
	'取主題館所有的nodeid
	sql = "SELECT CtNodeID FROM CatTreeNode WHERE CtRootID = '"&request("id")&"'"
	set rs2 = conn.execute(sql)
	while not rs2.eof
		'檢查原本有沒有權限
		sql_check = "select * from CtUserSet where UserID = '"& new_owner &"' and CtNodeID = '"&rs2("CtNodeID")&"'"
		set rs_check = conn.execute(sql_check)
		if not rs_check.eof then
			sql_del = "UPDATE CtUserSet SET Rights = '1' where UserID = '"& new_owner &"' and CtNodeID = '"&rs2("CtNodeID")&"'"
			conn.execute(sql_del)
		else
			sql_ins = "INSERT INTO CtUserSet (UserID,CtNodeID,Rights) VALUES ('"& new_owner &"','"&rs2("CtNodeID")&"','1')"
			conn.execute(sql_ins)
		end if
		rs2.movenext
	wend
	'response.end
	Response.Write "<script language='javascript'>"
	Response.Write "alert('主題館管理權限轉移成功！');"
	Response.Write "location.href = 'subject_set.asp?id="&request("id")&"';"
	Response.Write "</script>"
'上稿權限
elseif request("type") = 2 then
	'加入主題館上稿人員名單
	'更新主題館上稿人員名單
	if o_subject_user <> "" then
		o_users=split(o_subject_user,",")
		if request("ckbox") <>"" then
			new_user = request("ckbox")&","
		end if
		for i = 0 to Ubound(o_users)
			if not (instr(request("ckbox"),o_users(i))>0) then
				new_user = new_user & o_users(i) &","
			end if
		next
	else 
		new_user = request("ckbox") &","
	end if
	new_user = Left(new_user,Len(new_user)-1)
	sql_update = "UPDATE NodeInfo SET subject_user = '"& new_user &"' WHERE CtrootID = '"& request("id") &"'"
	conn.execute(sql_update)
	'增加新設定人員的上稿權限
	if request("ckbox") <> "" then
		users=split(request("ckbox"),",")
		for i = 0 to Ubound(users)
			'取主題館所有的nodeid
			sql = "SELECT CtNodeID FROM CatTreeNode WHERE CtRootID = '"&request("id")&"'"
			set rs2 = conn.execute(sql)
			while not rs2.eof
				'檢查原本有沒有權限
				sql_check = "select * from CtUserSet where UserID = '"&trim(users(i))&"' and CtNodeID = '"&rs2("CtNodeID")&"'"
				set rs_check = conn.execute(sql_check)
				if not rs_check.eof then
					sql_del = "UPDATE CtUserSet SET Rights = '1' where UserID = '"&trim(users(i))&"' and CtNodeID = '"&rs2("CtNodeID")&"'"
					conn.execute(sql_del)
				else
					sql_ins = "INSERT INTO CtUserSet (UserID,CtNodeID,Rights) VALUES ('"&trim(users(i))&"','"&rs2("CtNodeID")&"','1')"
					conn.execute(sql_ins)
				end if
				rs2.movenext
			wend
		next
	end if
	response.redirect "subject_set.asp?id="&request("id")

'刪除權限
elseif request("type") = 3 then
	'更新主題館上稿人員名單
	users=split(o_subject_user,",")
	new_user = ""
	for i = 0 to Ubound(users)
		if not (instr(users(i),request("ckbox"))>0) then
			new_user = new_user & users(i) &","
		end if
	next
	if new_user <> "" then
		new_user = Left(new_user,Len(new_user)-1)
	else
		new_user = new_user
	end if
	sql_update = "UPDATE NodeInfo SET subject_user = '"& new_user &"' WHERE CtrootID = '"& request("id") &"'"
	conn.execute(sql_update)
	'刪除上稿權限
	'取主題館所有的nodeid
	sql = "SELECT CtNodeID FROM CatTreeNode WHERE CtRootID = '"&request("id")&"'"
	set rs = conn.execute(sql)
	while not rs.eof
		sql_del = "UPDATE CtUserSet SET Rights = '0' where UserID = '"&trim(request("ckbox"))&"' and CtNodeID = '"&rs("CtNodeID")&"'"
		conn.execute(sql_del)
		rs.movenext
	wend
	'轉移上稿文章權限
	o_dept = request("deptID")
	'先檢查有無同單位上稿者
	Dim deptcheck
	deptcheck = False
	new_users=split(new_user,",")
	for i=0 to Ubound(new_users)
		sql_checkdept = "SELECT deptID FROM InfoUser WHERE UserID = '"& trim(new_users(i)) &"'"
		set re_checkdept = conn.execute(sql_checkdept)
		if o_dept = trim(re_checkdept("deptID")) and not re_checkdept.eof then
			deptcheck = True '有同單位上稿者
		end if
	next
	if not deptcheck then
		'沒有同單位上稿者 文章轉移到主題館管理者
		'取主題館所有的CtUnitID
		sql = "SELECT CtUnitID FROM CatTreeNode WHERE CtRootID = '"&request("id")&"'"
		set rs_ctunit = conn.execute(sql)
		while not rs_ctunit.eof
			if rs_ctunit("CtUnitID") <> "" then
				sql_update2 = "UPDATE CuDTGeneric SET iEditor = '"& o_owner &"' ,iDept ='"& o_owner_dept &"' WHERE iCTUnit = '"& rs_ctunit("CtUnitID") &"' and  iEditor ='"& request("ckbox") &"'"
				conn.execute(sql_update2)
			end if
			rs_ctunit.movenext
		wend
	end if
	Response.Write "<script language='javascript'>"
	Response.Write "alert('刪除成功！');"
	Response.Write "location.href = 'subject_set.asp?id="&request("id")&"';"
	Response.Write "</script>"
end if
response.end
%>