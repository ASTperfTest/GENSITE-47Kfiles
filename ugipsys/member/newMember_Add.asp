<%@ CodePage = 65001 %>
<% 
Response.Expires = 0
HTProgCap="單元資料維護"
HTProgFunc="編修"
HTUploadPath=session("Public")+"data/"
HTProgCode="PEDIA01"
HTProgPrefix="newMember" 
HTProgCode="newmember"
%>
<!--#include virtual = "/inc/client.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript" type="text/javascript">
<!--
function Send()
{
var rege = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9])+$/;

    if (document.Form1.account.value == ""  ) {
		alert("您忘了填寫帳號了！"); 
		document.Form1.account.focus(); 
	return false; 
	} 
	if(document.Form1.account.value.length > 30  ) {
		alert("您所填寫的帳號超過30碼！"); 
		document.Form1.account.focus(); 
	     return false; 
	 }
	if (document.Form1.account.value.length < 6){
        alert("您所填寫的帳號少於6碼！"); 
	    document.Form1.account.focus();
		return false; 
       }	
	if (document.Form1.account.value!=""){
       
	      var i;
		  var ch;
		  for (i=0;i< document.Form1.account.value.length;i++){
			ch = document.Form1.account.value.charAt(i);
			if(ch >= 'a' && ch <= 'z'){
				//return true;
			}
			else if(ch >= 'A' && ch <= 'Z'){
				//return true;
			}
			else if(ch >= '0' && ch <= '9'){
				//return true;
			}
			else if(ch == '-' || ch == '_'){
				//return true;
			}
			else{
				alert("帳號限用英文與數字，可用『-』或『_』！"); 
				return false;
			}
	      }
	}
	if (document.Form1.realname.value == ""  ) {
		alert("請輸入姓名！"); 
		document.Form1.realname.focus(); 
	return false; 
	}  
	if (document.Form1.passwd.value =="") {
	        alert("您忘了填寫密碼了！"); 
			document.Form1.passwd.focus(); 
			return false; 
		}
	if(document.Form1.passwd.value!=""){
            		
		            if (document.Form1.passwd.value.length > 16){
			            alert("您所填寫的密碼超過16碼！"); 
			            document.Form1.passwd.focus(); 
			             return false; 
		            }
		            else if(document.Form1.passwd.value.length < 6){
			            alert("您所填寫的密碼少於6碼！"); 
			            document.Form1.passwd.focus(); 
			            return false; 
		            }
		            else if(document.Form1.password2.value==""){
			            alert("您忘了填寫密碼確認了！"); 
			            document.Form1.password2.focus(); 
			            return false; 
		            }
		            else if(document.Form1.passwd.value != document.Form1.password2.value){
			            alert("密碼與密碼確認不符！");
			            document.Form1.password2.focus(); 
			            return false; 
		            }
		            else{
			            var i;
			            var ch;
			            var digits = "0123456789";
			            var digitflag = false;
			            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
			            var charflag = false;
			            for (i=0;i< document.Form1.passwd.value.length;i++){		
				            ch = document.Form1.passwd.value.charAt(i);		
				            if(ch == '\"' || ch == ' ' || ch == '\'' || ch== '\&'){
					            alert("密碼請勿包含『\"』、『'』、『&』或空白"); 
					            document.Form1.passwd.focus();
					            return false; 
				            }
				            if( digits.indexOf(ch) >= 0 ) {
					            digitflag = true;
				            }					
				            if( chars.indexOf(ch) >= 0 ) {
					            charflag = true;
				            }
			            }
			            if( !digitflag ) {
					            alert("密碼請至少包含一數字"); 
					            document.Form1.passwd.focus();
					            return false; 
			            }
			            if( !charflag ) {
					            alert("密碼請至少包含一英文字"); 
					           document.Form1.passwd.focus();
					           return false; 
			            }
		            }	
	
	
	}
	//if (document.Form1.id.value == "" ) {
	//	alert("請輸入身分證字號！"); 
	//	document.Form1.id.focus(); 
	//return false; 
	//}
	
	if(document.Form1.id_type2.checked)
	{
	
	if (document.Form1.member_org.value == ""  ) {
		alert("請輸入所屬機關名稱！"); 
		document.Form1.member_org.focus(); 
	return false; 
	}
	if (document.Form1.com_ext.value.length > 4  ) {
		alert("您所填寫所屬機關電話分機超過4碼！"); 
		document.Form1.com_ext.focus(); 
	return false; 
	} 
	
	 if (document.Form1.com_tel.value == ""  ) {
		alert("請輸入所屬機關電話！"); 
		document.Form1.com_tel.focus(); 
	return false; 
	} 
	//else if (document.Form1.com_exttext.value == ""  ) {
		//alert("請輸入所屬機關分機！"); 
		//document.Form1.com_exttext.focus(); 
	//return false; 
	//} 
	if (document.Form1.ptitle.value == ""  ) {
		alert("請輸入職稱！"); 
		document.Form1.ptitle.focus(); 
	return false; 
	} 
	}
	if (document.Form1.home_ext.value.length > 4  ) {
		alert("您所填寫電話分機超過4碼！"); 
		document.Form1.home_ext.focus(); 
	return false; 
	} 
	if (document.Form1.email.value == ""  ) {
		alert("請輸入E-mail！"); 
		document.Form1.email.focus(); 
	return false; 
	} 
	
	if (rege.exec(document.Form1.email.value) == null) {
		alert("eMail 格式錯誤！"); 
		document.Form1.email.focus(); 
		return false; 
	}
 
	
	//document.getElementsByName("submitTask").value = "UPDATE";
	//document.Form1.submit();
	 return true;
}

	function showbutton(id) {
		
		if (id == 1) {
			document.getElementById("orgnamediv").style.display = "none";	           	
			document.getElementById("orgtextdiv").style.display = "none";
            document.getElementById("com_telnamediv").style.display = "none";	           	
			document.getElementById("com_teltextdiv").style.display = "none";
			document.getElementById("com_extnamediv").style.display = "none";	           	
			document.getElementById("com_exttextdiv").style.display = "none";
			document.getElementById("ptitlenamediv").style.display = "none";	           	
			document.getElementById("ptitletextdiv").style.display = "none";
			document.getElementById("KMcatnamediv").style.display = "none";	           	
			document.getElementById("KMcattextdiv").style.display = "none";
		
		}
		if (id == 2) {
			document.getElementById("orgnamediv").style.display = "block";	           	
			document.getElementById("orgtextdiv").style.display = "block";	
            document.getElementById("com_telnamediv").style.display = "block";	           	
			document.getElementById("com_teltextdiv").style.display = "block";
	        document.getElementById("com_extnamediv").style.display = "block";	           	
			document.getElementById("com_exttextdiv").style.display = "block"; 
            document.getElementById("ptitlenamediv").style.display = "block";	           	
			document.getElementById("ptitletextdiv").style.display = "block"; 
            document.getElementById("KMcatnamediv").style.display = "block";	           	
			document.getElementById("KMcattextdiv").style.display = "block";         	
		}
	
  }  	
-->
</script>



<title>
	管理會員資料
</title>
<link href="../css/layout.css" rel="stylesheet" type="text/css" />
<link href="../css/form.css" rel="stylesheet" type="text/css" /></head>
<body>
    <div id="FuncName">
	    <h1>新增會員資料</h1>
		<div id="Nav">
	       <a href="javascript:history.go(-1)">回條列</a>
	      
		</div>
		<div id="ClearFloat"></div>
    </div>
    <div id="FormName">&nbsp;<font size="2">【新增會員資料】</font>    </div>
    <form name="Form1" method="post"  id="Form1" action="newMemberAdd_Act.asp" onsubmit="return Send()" ENCTYPE="MULTIPART/FORM-DATA" >
	 
<div>
</div>

         
          <table width="90%" cellspacing="0">
          <tr>
            <th colspan="2">會員資料管理</th>
            </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>會員帳號：</td>
            <td class="eTableContent"><input name="account" type="text"  id="account" value=""  />
			限用英文與數字，可用『-』或『_』，30碼以下</td>
          </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>真實姓名：</td>
            <td class="eTableContent"><input name="realname" type="text" id="realname" value="" /></td>
          </tr>
          <tr>
            <td class="Label" align="right">暱　　稱：</td>
            <td class="eTableContent"><input name="nickname" type="text" value="" id="nickname" /></td>
          </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>設定密碼：</td>
            <td class="eTableContent"><input name="passwd" type="password" id="passwd" value="" />
            
            自訂英文（區分大小寫）、數字，不含空白及@，16碼以下</td>
          </tr>
          <tr>
            <td class="Label" align="right"><span class="Must">*</span>密碼確認：</td>
            <td class="eTableContent"><input name="password2" type="password" id="password2" value="" /></td>
          </tr>                                    
          <tr>
            <td class="Label" align="right">身分證字號：</td>
            <td class="eTableContent"><input name="id" type="text"  id="id" value="" /></td>
          </tr>
		   <tr>
            <td class="Label" align="right">身分：</td>                  
            <td class="eTableContent" align="left"   >
             <input id="id_type1" type="radio" name="id_type" value="1" onclick="showbutton(1)" checked  /><label for="male">一般會員</label>
				
	         <input id="id_type2" type="radio" name="id_type"  value="2" onclick="showbutton(2)" />
				<label for="female">學者會員</label>	</td>
          </tr>
          <tr>
		    <td class="Label" align="right"><div id="orgnamediv" style="display:none"><span class="Must">*</span>所屬機關名稱：</div></td>
            <td class="eTableContent"><div id="orgtextdiv" style="display:none"><input name="member_org" type="text" value="" id="member_org" /></div></td>
		 </tr>
          <tr>
            <td class="Label" align="right"><div id="com_telnamediv" style="display:none"><span class="Must">*</span>所屬機關電話：</div></td> 
			<td class="eTableContent">
                <div id="com_teltextdiv" style="display:none"><input name="com_tel" type="text" value="" id="com_tel"  /></div></tr>
          <tr>
            <td class="Label" align="right"><div id="com_extnamediv" style="display:none"><label for="com_ext">分機：</label></div>
            <td><div id="com_exttextdiv" style="display:none"><input name="com_ext" type="text" id="com_ext" value="" /> </div></td>
				
          </tr>
          <tr>
            <td class="Label" align="right"> <div id="ptitlenamediv" style="display:none"><span class="Must">*</span>職稱：</div></td>
            <td class="eTableContent"><div id="ptitletextdiv" style="display:none"><input name="ptitle" type="text" value="" id="ptitle" /></div></td>
          </tr>
          <tr>
            <td class="Label" align="right">出生日期：</td>      
            <td class="eTableContent">

        西元<input name="birthYear" type="text" value="" id="birthYear" />年
				
       <select name="birthMonth" id="birthMonth">
    <option value="" selected>請選擇</option>
	<option value="01" >1</option>
	<option value="02" >2</option>
	<option value="03" >3</option>
	<option value="04" >4</option>
	<option value="05" >5</option>
	<option value="06" >6</option>
	<option value="07" >7</option>
	<option value="08" >8</option>
	<option value="09" >9</option>
	<option value="10" >10</option>
	<option value="11" >11</option>
	<option  value="12" >12</option>
</select>
                月
                <select name="birthday" id="birthday">
	<option value="" selected>請選擇</option>	
	<option value="01" >1</option>
	<option value="02" >2</option>
	<option value="03" >3</option>
	<option value="04" >4</option>
	<option value="05" >5</option>
	<option value="06" >6</option>
	<option value="07" >7</option>
	<option value="08" >8</option>
	<option value="09" >9</option>
	<option value="10" >10</option>
	<option value="11" >11</option>
	<option value="12" >12</option>
	<option value="13" >13</option>
	<option value="14" >14</option>
	<option value="15" >15</option>
	<option value="16" >16</option>
	<option value="17" >17</option>
	<option value="18" >18</option>
	<option value="19" >19</option>
	<option value="20" >20</option>
	<option value="21" >21</option>
	<option value="22" >22</option>
	<option value="23" >23</option>
	<option value="24" >24</option>
	<option value="25" >25</option>
	<option value="26" >26</option>
	<option value="27" >27</option>
	<option value="28" >28</option>
	<option value="29" >29</option>
	<option value="30" >30</option>
	<option value="31" >31</option>
</select>       
                日			</td>
          </tr>
		
          <tr>
            <td class="Label" align="right">性別：</td>                  
            <td class="eTableContent" align="left">
             <input id="male" type="radio" name="sex"value="1"  checked  /><label for="male">男</label>
				
	<input id="female" type="radio" name="sex"  value="0"  />
				<label for="female">女</label>	</td>
          </tr>
		  
		
          <tr>
            <td class="Label" align="right">地址：</td>               
            <td class="eTableContent">
			    <input name="homeaddr" type="text" value=""  id="homeaddr" />
                <label for="zip"> 郵遞區號：</label>
                <input name="zip" type="text" id="zip" value="" />    </td>
          </tr>
          <tr>
            <td class="Label" align="right">電話：</td>              
            <td class="eTableContent">
              <input name="phone" type="text" value=""                                                                                                     " id="phonetext" />
              <label for="home_ext">分機：</label>
              <input name="home_ext" type="text" id="home_ext" value=""  />            </td>
          </tr>
          <tr>
            <td class="Label" align="right">手機號碼：</td>    
            <td class="eTableContent"><input name="mobile" type="text" value=""                                                                                                  " id="mobiletext" /></td>
          </tr>
          <tr>            
            <td class="Label" align="right">傳真：</td>    
            <td class="eTableContent"><input name="fax" type="text" id="fax" value="" /></td>
          </tr>
          <tr>            
            <td class="Label" align="right"><span class="Must">*</span>E-mail：</td>    
            <td class="eTableContent"><input name="email" type="text" value="" id="email" />
			供系統認證之用，請務必填寫正確</td>
          </tr>
		  <tr>
			   <td class="Label" align="right">備註：</td>	
			   <td><textarea name="remark" cols="50" rows="4" id="remark" class="InputText"></textarea></td>
		  </tr>
          <tr>
            <td class="Label" align="right"><div id="KMcatnamediv" style="display:none">研究領域及專長：</div></td>
            <td> <div id="KMcattextdiv" style="display:none"><input type="text" size="30" name="KMcat"  id="KMcat" value="" ></div>
						<input class="Text" type="hidden" size="60" name="htx_KMcatID"  id="htx_KMcatID" >
						<input class="Text" type="hidden" size="60" name="htx_KMautoID"  id="htx_KMautoID" >
					
			
                            </td>
          </tr>
          </table>
          <hr />
          <table width="90%" cellspacing="0">

            <tr>
              <th colspan="2">會員權限管理</th>
            </tr>
            <tr>
              <td class="Label" align="right">上傳圖片權限：</td>
              <td class="eTableContent"> 目前狀態為
                <!--input type="submit" name="Cat" value="開啟" onclick="GetCat();" id="Cat" class="cbutton" /-->
				
                  <select name="uploadRight" size="1">
                    <option value=""  selected >請選擇</option>
                    <option value="Y"  >正常</option>
                    <option value="N"  >停權</option>
                </select>
				可上傳圖片數量/ 每次
				<input name="uploadPicCount" type="text" id="uploadPicCount" value="1" size="10" />
				張</td>
            </tr>
            <!--<tr>
              <td align="right" class="Label">會員身分管理：</td>
              <td class="eTableContent">一般會員 / 學者會員
                <select name="scholarValidate" size="1">
				<option value="" selected>請選擇</option>	
                    <option value="Z" >無</option>
                    <option value="W"  >待審核</option>
                    <option value="Y"  >通過</option>
                    <option value="N"  >不通過</option>
                </select></td>
            </tr>
			-->
          </table>
          <hr />
          <table width="90%" cellspacing="0">

            <tr>
              <th colspan="2">專家身分管理</th>
            </tr>
            <tr>
              <td class="Label" align="right">專家：</td>
              <td class="eTableContent"><input name="id_type3" type="checkbox" value="1"    />
                  <span class="Label">設為知識家專家</span></td>
            </tr>
						<!--
						<tr>
							<td class="Label" align="right">專長關鍵字：</td>
							<td class="eTableContent">
								<input name="htx_xkeyword3" class="rdonly" title="請以,分隔" value="" size="75" readonly="true" /><br/>
								<input name="button2" type="button" class="cbutton" onclick="MM_goToURL('self','keywords_set.htm');return document.MM_returnValue" value ="設定" />              
							</td>
						</tr>
						-->
            <tr>
              <td class="Label" align="right">專家圖檔：</td>
			  
              <td class="eTableContent"><input type="file" name="photo" id="photo" /></td>
			  
            </tr>
          </table>

		   <input type="submit" name="submit" value ="編修存檔" " class="cbutton" >
		  
            
           <input type="button" name="ResetBtn" class="cbutton" value="取消" onclick="returnToEdit()" id="ResetBtn" class="cbutton" />   
           
           

        
    
<div>

</div>
</form>
</body>
</html>
<script language="javascript">
	
	function returnToEdit() {
		window.location.href = "newMemberList.asp";
     }   
    
</script>



