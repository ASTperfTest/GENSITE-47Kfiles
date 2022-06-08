<%@  codepage="65001" %>
<% 
Response.Expires = 0
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.CacheControl = "Private"
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="newmember"
HTProgPrefix = "newMember"
%>
<!--#include virtual = "/inc/server.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <link href="/css/list.css" rel="stylesheet" type="text/css">
    <link href="/css/layout.css" rel="stylesheet" type="text/css">
    <title></title>
</head>
<body>
    <form id="Form" name="reg" id="reg" method="post" action="newMemberQuery_Act.asp?request=yes">
        <input type="hidden" name="submitTask" id="submitTask" value="">
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td width="50%" align="left" nowrap class="FormName">
                    會員整合管理&nbsp; <font size="2">【查詢】</font>
                </td>
                <td width="50%" class="FormLink" align="right">
                    <a href="Javascript:window.history.back();" title="回前頁">回前頁</a>
                </td>
            </tr>
            <tr>
                <td width="100%" colspan="2">
                    <hr noshade size="1" color="#000080">
                </td>
            </tr>
            <tr>
                <td class="Formtext" colspan="2" height="15">
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" width="80%" height="230" valign="top">
                    <center>
                        <table border="0" id="ListTable" cellspacing="1" cellpadding="2">
                            <tr>
                                <th>帳號：</th>
                                <td class="eTableContent">
                                    <input name="account" size="50" value="">
                                </td>
                            </tr>
                            <tr>
                                <th>真實姓名：</th>
                                <td class="eTableContent">
                                    <input name="realname" size="50" value="">
                                </td>
                            </tr>
                            <tr>
                                <th>暱稱：</th>
                                <td class="eTableContent">
                                    <input name="nickname" size="50" value="">
                                </td>
                            </tr>
                            <tr>
                                <th>電子郵件：</th>
                                <td class="eTableContent">
                                    <input name="emailaddress" size="50" value="">
                                </td>
                            </tr>
							<tr>
                                <th>身份字號：</th>
                                <td class="eTableContent">
                                    <input name="idcard" size="50" value="">
                                </td>
                            </tr>
                            <tr>
                                <th>地址：</th>
                                <td class="eTableContent">
                                    <input name="homeaddr" size="50" value="">
                                </td>
                            </tr>
                            <tr>
                                <th>研究領域及專長：</th>
                                <td class="eTableContent">
                                    <input name="KMcat" type="text" id="KMcat" />
                                    <input type="hidden" name="htx_KMcatIDtext" id="htx_KMcatIDtext" />
                                    <input type="hidden" name="htx_KMautoIDtext" id="htx_KMautoIDtext" />
                                    <input type="hidden" name="Cat" value="外部分類" onclick="GetCat();" id="Cat" class="cbutton" />
                                </td>
                                <tr>
                                    <th>專長關鍵字：</th>
                                    <td class="eTableContent">
                                        <input name="keyword" size="50">
                                    </td>
                                </tr>
                            <tr>
                                    <th>備註：</th>
                                    <td class="eTableContent">
                                        <input name="remark" size="50">
                                    </td>
                                </tr>
                            <tr>
                                <th>會員身分：</th>
                                <td class="eTableContent">
                                    <input name="id_type1" type="checkbox" value="1" />
                                    一般會員
                                    <input name="id_type2" type="checkbox" value="1" />
                                    學者會員
                                </td>
                            </tr>
                            <tr>
                                <th>專家身分：</th>
                                <td class="eTableContent">
                                    <input id="id_type3" type="radio" onclick='kmintra.disabled=false;' name="id_type3" value="1" />
                                    <label for="maletext">是</label>
                                    <input id="id_type30" type="radio" onclick='kmintra.checked=false; kmintra.disabled=true;' name="id_type3" value="0" />
                                    <label for="femaletext">否</label>
									<!--Grace-->
									<input id="kmintra" name="kmintra" type="checkbox"  value="1" disabled = "true" />
									是否包含kmintra
									
                                </td>
                            </tr>
                            <tr>
                                <th>學者審核狀態</th>
                                <td class="eTableContent">
                                    <input name="scholarValidate" type="radio" value="W" />
                                    待審核
                                    <input name="scholarValidate" type="radio" value="Y" />
                                    通過
                                    <input name="scholarValidate" type="radio" value="N" />
                                    不通過
                                    <input name="scholarValidate" type="radio" value="Z" />
                                    不需審核
                                </td>
                            </tr>
                            <tr>
                                <th>會員狀態：</th>
                                <td class="eTableContent">
                                    <input name="status" type="radio" value="Y" />正常
                                    <input name="status" type="radio" value="N" />停權
                                </td>
                            </tr>                            
                            <tr>
                                <th>Email驗證狀態：</th>
                                <td class="eTableContent">
                                    <input name="mcode" type="radio" value="Y" />已驗證<br/>
                                    <input name="mcode" type="radio" value="Sended" />已申請驗證 &nbsp
                                    <input name="mcode" type="radio" value="Apply" />未申請驗證 &nbsp
                                </td>
                            </tr>
                        </table>
                </td>
            </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
            <tr>
                <td width="100%">
                    <p align="center">
                        <input name="button" type="button" class="cbutton" onclick="QuerySubmit()" value="查 詢">
                        <input type="button" value="重　填" class="cbutton" onclick="returnToEdit()">
                </td>
            </tr>
        </table>
        </TABLE>
        <!-- 程式結束 ---------------------------------->
    </form>

    <script type="text/javascript">
        function QuerySubmit() {
            //document.getElementById("picId").value = id;
            //document.getElementById("parentIcuitem").value = parentIcuitem;
            document.getElementById("submitTask").value = "Query";
            //		document.getElementById("reg")
            document.forms[0].submit();
        }
        function returnToEdit() {
            //window.location.href = "cp_question.asp?iCUItem=<%=request.querystring("iCUItem")%>&nowPage=<%=request.querystring("nowPage")%>&pagesize=<%=cint(request.querystring("pagesize"))%>";
            window.location.href = "newMember_Query.asp";

        }
	
    </script>

</body>
</html>
