<!--#Include file = "../inc/server.inc" -->
<!--#INCLUDE FILE="../inc/dbutil.inc" -->
<%
if request.form("doJob")<>"" then        '½T©w°õ¦æ®É
  for each x in request.form
	if left(x,5)="ckbox" and request(x)<>"" then     
		xn=mid(x,6)
