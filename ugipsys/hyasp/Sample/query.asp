<html>
<title>Hysearch �����˯����� ASP ���սd��</title>
<body>
<center>
<br>
<h3>HySearch �����˯��t��  ASP �������պ���</h3><br>
�d�ߵ��G<br><br>
<%
   ' step 1: �ϥ� CreateObject ���� Object. DLL name: hysdk, class name: hyft
   Set obj = Server.CreateObject("hysdk.hyft.1")

   ' step 2: �I�s hyft_connect �s�� hyftd
   x = obj.hyft_connect("localhost", 2816, "WRITER", "", "")

   ' hyft_connect() �Ǧ^�Ȭ��s�u id, �s��� obj.ret_val ��
   id = obj.ret_val
   if id = -1 then
      '�Y�O�s�u����, �h���~�T���s��� obj.errmsg
      response.write "�s�u hyftd �o�Ϳ��~, ���~�T��: " & obj.errmsg & "<br><br>"
      response.end
   else
      response.write "�s�u hyftd ���\! id = " & id & "<br><br>"
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

   ' step 4: �}�l�I�s hyft_addquery, �N�n�d�ߤ�����ƶ�J
   response.write "�d�߱���Gau=�k���J�L<br><br>"
   x = obj.hyft_addquery(id, "au", "�k���J�L", "", 0, 0, 0)
   if ret = -1 then
      response.write "addquery �o�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if

   ' step 5: �I�s hyft_query, �}�l�d��
   x = obj.hyft_query(id, 0, 1000)
   res_id = obj.ret_val
   if res_id = -1 then
      response.write "query �o�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if 

   ' step 6: ���o����
   x = obj.hyft_num_sysid(id, res_id)
   ' ���ƶǦ^�Ȧs��� obj.ret_val
   num = obj.ret_val
   if num < 0 then
      response.write "num_sysid �o�Ϳ��~, ���~�T��: " & obj.errmsg
      response.end
   end if 
   ' num > 0 ��ܦ��d����
   if num > 0 then
      ' step 6-1: �@������@�� sysid
      for i = 0 to num-1
	x = obj.hyft_fetch_sysid(id, res_id, i)
        '����X�Ӥ� sysid �s��� obj.sysid
        sysid = obj.sysid
        response.write "sysid=***" & sysid & "***<br><br>"
      next
      response.write "����!"
   else
      response.write "num = 0, no data found."
   end if
   set obj=Nothing	
%>
<br><br>
<a href="javascript:history.go(-1)">�^�W��</a>
</center>
</body>
</html>
