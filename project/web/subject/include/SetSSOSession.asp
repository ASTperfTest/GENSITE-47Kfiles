
<%
function SetSSOSession(ssoId)

    session("backEndSetup") = ""
    ssoId = replace(ssoId, "'", "aaaaaaaaaaaaaaaa")

    sqlstr = "select * from sso "
    sqlstr = sqlstr & vbcrlf & "where [user_id] like 'uAdm_%' and (creation_datetime>getdate()-1) "
    sqlstr = sqlstr & vbcrlf & "and [guid] = '" & ssoId & "'"
    
    if ssoId <>"" then
        sqlstr = "select * from sso "
        sqlstr = sqlstr & vbcrlf & "where [user_id] like 'uAdm_%' and (creation_datetime>getdate()-1) "
        sqlstr = sqlstr & vbcrlf & "and [guid] = '" & ssoId & "'"
        set rs = conn.execute(sqlstr)		    
        if not rs.eof then
            session("backEndSetup") = ssoId
        end if
    end if
    
end function
%>