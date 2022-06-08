<%@ CodePage = 65001 %>
<!--#include virtual = "/inc/client.inc" -->
<!--#include virtual = "/inc/dbFunc.inc" -->
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"; cache=no-caches;>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../inc/setstyle.css">
<title>查詢結果清單</title></head><body bgcolor="#FFFFFF" background="../images/management/gridbg.gif">
<%
  '取得輸入
  submit = trim(request("submit"))
  account = trim(request("account"))

  '確定
  if trim(submit) = "delete" then

    sql = "delete from member where account=N'" & account & "'"
    conn.execute(sql)

    '導回首頁
    response.write "<script language='JavaScript'>alert('刪除會員資料完成！');location.replace('memberList.asp');</script>"
    response.end

  end if

  if trim(submit) = "確定" then
  
    '取得輸入
    passwd1 = request("passwd1")
    passwd2 = request("passwd2")
    realname = request("realname")
    homeaddr = request("homeaddr")
    phone = request("phone")
    mobile = request("mobile")
    email = request("email")
    mcode = request("mcode")
    
    '檢查輸入
    if trim(passwd1) = "" then
      response.write "<script language='JavaScript'>alert('請輸入密碼！');history.go(-1);</script>"
      response.end
    end if
    if trim(passwd2) = "" then
      response.write "<script language='JavaScript'>alert('請輸入密碼確認！');history.go(-1);</script>"
      response.end
    end if
    if trim(realname) = "" then
      response.write "<script language='JavaScript'>alert('請輸入姓名！');history.go(-1);</script>"
      response.end
    end if
    if trim(email) = "" then
      response.write "<script language='JavaScript'>alert('請輸入電子信箱！');history.go(-1);</script>"
      response.end
    end if
    if trim(mcode) = "" then
      response.write "<script language='JavaScript'>alert('請輸入身分群組！');history.go(-1);</script>"
      response.end
    end if
    
    '檢查密碼
    if trim(passwd1) <> trim(passwd2) then
      response.write "<script language='JavaScript'>alert('兩次密碼不相同！');history.go(-1);</script>"
      response.end
    end if
    
    '檢查email格式是否正確
    if len(email) > 3 and InStr(email, "@") > 0 and InStr(email, ".") > 0 then
      '正確
    else  
      '不正確...跳回
      response.write "<script language='JavaScript'>alert('您的電子信箱格式輸入錯誤！');history.go(-1);</script>"
      response.end
    end if
       
    '存入資料
    sql = "update Member set passwd='" & passwd1 & "',realname='" & realname & "',homeaddr='" & homeaddr & "',phone='" & phone & "',mobile='" & mobile & "',email='" & email & "',mcode='" & mcode & "',modifytime=getdate() where account='" & account & "'"
    conn.execute(sql)
    
    '導回首頁
    response.write "<script language='JavaScript'>alert('修改會員資料完成！');location.replace('memberList.asp');</script>"
    response.end
    
  end if

  '導回首頁
  response.write "<script language='JavaScript'>location.replace('memberList.asp');</script>"
  response.end
%>
</body>
</html>