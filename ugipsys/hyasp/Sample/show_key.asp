
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<title>
Hyftd_asp
</title>
</head>

<body>
<h1>
Hyftd asp_demo
</h1>
<form method="post" action="show_key.asp">
<br>Enter new value   :  <input name="query" type="text"><br>
<br><br>
<input type="submit" name="Submit" value="Submit">
<input type="reset" value="Reset">
<br>

<%
   
   ' step 1: �ϥ� CreateObject ���� Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   ' step 2: �I�s hyft_connect �s�� hyftd	hyft_connect() �Ǧ^�Ȭ��s�u id, �s��� obj.ret_val ��
   	x = obj.hyft_connect("localhost", 2816, "gg", "", "")
	id = obj.ret_val
  
   if id = -1 then
      '�Y�O�s�u����, �h���~�T���s��� obj.errmsg
      response.write "�s�u hyftd �o�Ϳ��~, ���~�T��: " & obj.errmsg & "<br><br>"
      response.end
   end if

   query = Request("query")
 	
	' step 3: �ҥ�hyft_textkeys�_���\��,�p�G���\,�|�o��@��key_id,
	x=obj.hyft_textkeys(id, query,10)
	key_id = obj.ret_val
	if key_id = -1 then
      '�Y�O�s�u����, �h���~�T���s��� obj.errmsg
      response.write "�s�u hyftd �o�Ϳ��~, ���~�T��: " & obj.errmsg & "<br><br>"
      response.end
    end if
	
	
	' step 4: ��Xquery���_�����ƥ�
	x=obj.hyft_num_keys(id,key_id)
	num =obj.ret_val
	response.write "�@��" & num & "���_��<br><br>"
	
	for i = 0 to num-1
		x=obj.hyft_keylen(id,key_id,i)
		key_len=obj.ret_val
			
		response.write "key_len:" & key_len & "<br><br>"
		x=obj.hyft_onekey(id,key_id,i,key_len)
		'response.write "���~�T��: " & obj.errmsg & "<br><br>"
		key=obj.key
		response.write "key:" & key & "<br><br>"
		
	next 
	
	
	x=obj.hyft_free_key(id,key_id)
	
   	set obj=Nothing	
%>
</form>
</body>
</html>
