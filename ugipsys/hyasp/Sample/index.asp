
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
<form method="post">
<br>Enter new value   :  <input name="query" type="text"><br>
<br><br>
<input type="submit" name="Submit" value="Submit">
<input type="reset" value="Reset">
<br>

<%
   
   gg="gg"
   ' step 1: �ϥ� CreateObject ���� Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   ' step 2: �I�s hyft_connect �s�� hyftd
   x = obj.hyft_connect("localhost", 2816, "gg", "", "")

   ' hyft_connect() �Ǧ^�Ȭ��s�u id, �s��� obj.ret_val ��
   id = obj.ret_val
  
   if id = -1 then
      '�Y�O�s�u����, �h���~�T���s��� obj.errmsg
      response.write "�s�u hyftd �o�Ϳ��~, ���~�T��: " & obj.errmsg & "<br><br>"
      response.end
   else
      'response.write "�s�u hyftd ���\! id = " & id & "<br><br>"
   end if

   ' step 3: ��l�Ƭd������
   x = obj.hyft_initquery(id)
   ' �Ǧ^�Ȭ� obj.ret_val
   ret = obj.ret_val 
   if ret = -1 then
      '�Y�O��l�ƥ���, �h���~�T���s��� obj.errmsg
      response.write "initquery �o�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if
   
   query = Request("query")
   
   x = obj.hyft_addquery(id, "", query, "", 0, 0, 0)   
   ret = obj.ret_val
   if ret = -1 then
      response.write "addquery �o�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if

   'step 5: �I�s hyft_query, �}�l�d��(���Ƭ�1��1000)
    x = obj.hyft_query(id, 0,1000)
 	res_id= obj.ret_val
 	'response.write "res_id = " & res_id & "<br><br>"
   	
   	' step 6: ���o����,���ƶǦ^�Ȧs��� obj.ret_val
   	x = obj.hyft_num_sysid(id,res_id)
   	num = obj.ret_val
   	
   	' step 7: ���o������ƪ�id
   	x = obj.hyft_get_assdata_by_resultset(id,res_id,"gg")
   	assdata_id = obj.ret_val 
   	
   	' step 8: ���o���ޤά�����Ƥ��e
   	if num >0 then
   		for i = 0 to num-1
			
			'��X�C�@����ƪ�SYSID
			x = obj.hyft_fetch_sysid(id, res_id, i)
        	sysid = obj.sysid
        	'response.write "sysid=***" & sysid & "***<br>"
        	
        	'��X�C�@����ƪ�������ƪ���
        	x=obj.hyft_one_assdata_len(id,gg, sysid)
        	data_len=obj.ret_val
        	'response.write "data_len=***" & data_len& "***<br>"
        	
        	'��X�C�@����ƪ�������Ƥ��e
        	x=obj.hyft_one_assdata(id,data_len)
       		data =obj.data   
       		response.write "data=***" &  data & "***<br>"
    		
    	next 
    	
    	'����������id
    	x=obj.hyft_free_assdata(id, assdata_id)
   	else
   		response.write "�d�L���"
   	end if
   	
   set obj=Nothing	
%>
</form>
</body>
</html>
