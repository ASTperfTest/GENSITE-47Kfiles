<!-- #include file = "exam.config.asp" -->
<!-- #include file = "exam_util.inc.asp" -->
<%
Set connHakka = Server.CreateObject("ADODB.Connection")

'----------HyWeb GIP DB CONNECTION PATCH----------
'connHakka.Open Session("ODBCDSN")
'Set connHakka = Server.CreateObject("HyWebDB3.dbExecute")
connHakka.ConnectionString = Session("ODBCDSN")
connHakka.ConnectionTimeout=0
connHakka.CursorLocation = 3
connHakka.open
'----------HyWeb GIP DB CONNECTION PATCH----------

%>
