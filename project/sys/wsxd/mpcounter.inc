<%	
	sql = "SELECT * FROM counter where mp='" & request("mp") & "'"
	Set rs = conn.Execute(sql)
	If Not rs.EOF Then
		count = rs("counts") + 1
		sql = "UPDATE counter SET counts = counts + 1  WHERE mp='" & request("mp") & "'"
	Else
		count = 1
		sql="INSERT INTO counter (mp, counts) VALUES ('" & request("mp") & "','1')"
	End If
	Response.Write "<Counter>" & count & "</Counter>"
	Set rs = conn.Execute(sql)
%>