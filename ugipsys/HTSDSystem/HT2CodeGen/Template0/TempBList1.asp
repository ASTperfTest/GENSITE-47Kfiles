<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
if request.form("doJob")<>"" then        '�T�w�����
  for each x in request.form
	if left(x,5)="ckbox" and request(x)<>"" then     
		xn=mid(x,6)
