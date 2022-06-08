<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ProsDataAdd.aspx.vb" Inherits="ProsData_ProsDataAdd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>未命名頁面</title>
    <link href="/css/layout.css" rel="stylesheet" type="text/css"/>
    <link href="/css/form.css" rel="stylesheet" type="text/css"/>
</head>
<body>
    <div id="FuncName">
	    <h1>專家資料／新增專家資料</h1><font size="2">【目錄樹節點: 新增專家資料】</font>
	    <div id="Nav"></div>
	    <div id="ClearFloat"></div>
    </div>
    <div id="FormName">		       
       	單元資料維護&nbsp;<font size="2">【編輯(<font color="red">一般資料式</font>)--主題單元:新增專家資料 / 單元資料:純網頁】</font>
    </div>
    <form id="form1" runat="server" name="form1" enctype="multipart/form-data">
        <div>    
          <table  cellspacing="0">
          <tr>
            <td class="Label" align="right">*會員帳號：</td>
            <td class="eTableContent"><asp:TextBox ID="accounttext" runat="server"></asp:TextBox>限用英文與數字，可用『-』或『_』，30碼以下</td>
          </tr>
          <tr>
            <td class="Label" align="right">*真實姓名：</td>
            <td class="eTableContent"><asp:TextBox ID="realnametext" runat="server"></asp:TextBox></td>
          </tr>
          <tr>
            <td class="Label" align="right">暱　　稱：</td>
            <td class="eTableContent"><asp:TextBox ID="nicknametext" runat="server"></asp:TextBox></td>
          </tr>
          <tr>
            <td class="Label" align="right">*設定密碼：</td>
            <td class="eTableContent"><asp:TextBox ID="password1text" runat="server" TextMode="Password"></asp:TextBox>自訂英文（區分大小寫）、數字，不含空白及@，16碼以下</td>
          </tr>
          <tr>
            <td class="Label" align="right">*密碼確認：</td>
            <td class="eTableContent"><asp:TextBox ID="password2text" runat="server" TextMode="password"></asp:TextBox></td>
          </tr>                                    
          <tr>
            <td class="Label" align="right">*身分證字號：</td>
            <td class="eTableContent"><asp:TextBox ID="idntext" runat="server"></asp:TextBox></td>
          </tr>
          <tr>
            <td class="Label" align="right">*所屬機關名稱：</td>
            <td class="eTableContent"><asp:TextBox ID="member_orgtext" runat="server"></asp:TextBox></td>
          </tr>
          <tr>
            <td class="Label" align="right">*所屬機關電話：</td>
            <td class="eTableContent">
                <asp:TextBox ID="com_teltext" runat="server"></asp:TextBox>
                <label for="com_ext">分機：</label>
                <asp:TextBox ID="com_exttext" runat="server"></asp:TextBox>
            </td>
          </tr>
          <tr>
            <td class="Label" align="right">*職稱：</td>
            <td class="eTableContent"><asp:TextBox ID="ptitletext" runat="server"></asp:TextBox></td>
          </tr>
          <tr>
            <td class="Label" align="right">出生日期：</td>      
            <td class="eTableContent">
                西元<asp:TextBox ID="birthYeartext" runat="server"></asp:TextBox>年
                <asp:DropDownList ID="birthMonthtext" runat="server">
                    <asp:ListItem Text="1" Value="01"></asp:ListItem><asp:ListItem Text="2" Value="02"></asp:ListItem><asp:ListItem Text="3" Value="03"></asp:ListItem>
                    <asp:ListItem Text="4" Value="04"></asp:ListItem><asp:ListItem Text="5" Value="05"></asp:ListItem><asp:ListItem Text="6" Value="06"></asp:ListItem>
                    <asp:ListItem Text="7" Value="07"></asp:ListItem><asp:ListItem Text="8" Value="08"></asp:ListItem><asp:ListItem Text="9" Value="09"></asp:ListItem>
                    <asp:ListItem Text="10" Value="10"></asp:ListItem><asp:ListItem Text="11" Value="11"></asp:ListItem><asp:ListItem Text="12" Value="12"></asp:ListItem>
                </asp:DropDownList>
                月
                <asp:DropDownList ID="birthdaytext" runat="server">
                    <asp:ListItem Text="1" Value="01"></asp:ListItem><asp:ListItem Text="2" Value="02"></asp:ListItem><asp:ListItem Text="3" Value="03"></asp:ListItem>
                    <asp:ListItem Text="4" Value="04"></asp:ListItem><asp:ListItem Text="5" Value="05"></asp:ListItem><asp:ListItem Text="6" Value="06"></asp:ListItem>
                    <asp:ListItem Text="7" Value="07"></asp:ListItem><asp:ListItem Text="8" Value="08"></asp:ListItem><asp:ListItem Text="9" Value="09"></asp:ListItem>
                    <asp:ListItem Text="10" Value="10"></asp:ListItem><asp:ListItem Text="11" Value="11"></asp:ListItem><asp:ListItem Text="12" Value="12"></asp:ListItem>
                    <asp:ListItem Text="13" Value="13"></asp:ListItem><asp:ListItem Text="14" Value="14"></asp:ListItem><asp:ListItem Text="15" Value="15"></asp:ListItem>
                    <asp:ListItem Text="16" Value="16"></asp:ListItem><asp:ListItem Text="17" Value="17"></asp:ListItem><asp:ListItem Text="18" Value="18"></asp:ListItem>
                    <asp:ListItem Text="19" Value="19"></asp:ListItem><asp:ListItem Text="20" Value="20"></asp:ListItem><asp:ListItem Text="21" Value="21"></asp:ListItem>
                    <asp:ListItem Text="22" Value="22"></asp:ListItem><asp:ListItem Text="23" Value="23"></asp:ListItem><asp:ListItem Text="24" Value="24"></asp:ListItem>
                    <asp:ListItem Text="25" Value="25"></asp:ListItem><asp:ListItem Text="26" Value="26"></asp:ListItem><asp:ListItem Text="27" Value="27"></asp:ListItem>
                    <asp:ListItem Text="28" Value="28"></asp:ListItem><asp:ListItem Text="29" Value="29"></asp:ListItem><asp:ListItem Text="30" Value="30"></asp:ListItem>
                    <asp:ListItem Text="31" Value="31"></asp:ListItem>
                </asp:DropDownList>       
                日
			</td>
          </tr>
          <tr>
            <td class="Label" align="right">性別：</td>                  
            <td class="eTableContent">
                <asp:RadioButton ID="maletext" Text="男" runat="server" /><asp:RadioButton ID="femaletext" Text="女" runat="server" />
			</td>
          </tr>
          <tr>
            <td class="Label" align="right">地址：</td>               
            <td class="eTableContent">
			    <asp:TextBox ID="homeaddrtext" runat="server"></asp:TextBox>
                <label for="zip"> 郵遞區號：</label>
                <asp:TextBox ID="ziptext" runat="server"></asp:TextBox>
            </td>
          </tr>
          <tr>
            <td class="Label" align="right">電話：</td>              
            <td class="eTableContent">
              <asp:TextBox ID="phonetext" runat="server"></asp:TextBox>
              <label for="home_ext">分機：</label>
              <asp:TextBox ID="home_exttext" runat="server"></asp:TextBox>
            </td>
          </tr>
          <tr>
            <td class="Label" align="right">手機號碼：</td>    
            <td class="eTableContent"><asp:TextBox ID="mobiletext" runat="server"></asp:TextBox></td>
          </tr>
          <tr>            
            <td class="Label" align="right">傳真：</td>    
            <td class="eTableContent"><asp:TextBox ID="faxtext" runat="server"></asp:TextBox></td>
          </tr>
          <tr>            
            <td class="Label" align="right">圖檔：</td>    
            <td class="eTableContent"><asp:FileUpload ID="imgfile" runat="server" /></td>
          </tr>
          <tr>            
            <td class="Label" align="right">*E-mail：</td>    
            <td class="eTableContent"><asp:TextBox ID="emailtext" runat="server"></asp:TextBox>供系統認證之用，請務必填寫正確</td>
          </tr>
          <tr>            
            <td class="Label" align="right">研究領域及專長：</td>    
			<td class="eTableContent">
			    <asp:TextBox ID="htx_KMcat" runat="server"></asp:TextBox>
                <asp:HiddenField ID="htx_KMcatID" runat="server" />
                <asp:HiddenField ID="htx_KMautoID" runat="server" />
                <asp:Button ID="Cat" CssClass="cbutton" runat="server" Text="外部分類" OnClientClick="GetCat();"  />				
			</td>
          </tr>
          </table>
          <script language="javascript" type="text/javascript">
            function GetCat() {
                document.domain = "coa.gov.tw";
                window.open("http://kmintra.coa.gov.tw/coa/ekm/manage_doc/report2cat22.jsp?data_base_id=DB020&id_name=htx_KMcatID&autoid_name=htx_KMautoID&nm_name=htx_KMcat&&subNode=1*RB1&display=1*none&form_name=Form1&focus=&catidsInput=&anchor=",null,"height=400,width=500,status=no,scrollbars=yes,toolbar=no,menubar=no,location=no");
            }
          </script>
           <asp:Button ID="SubmitBtn" CssClass="cbutton" runat="server" Text="確定送出" OnClientClick="send();" />
           <asp:Button ID="ResetBtn" CssClass="cbutton" runat="server" Text="重新填寫" OnClientClick="document.Form1.reset();"/>              
           <script language="JavaScript">            
            function send(){
	            var form=document.Form1; 
	            
	            if(form.accounttext.value==""){
		            alert("您忘了填寫帳號了！"); 
		            form.accounttext.focus();
		            event.returnValue = false;
		            return;		            
	            }	
	            else{
		            if (form.accounttext.value.length > 30){
			            alert("您所填寫的帳號超過30碼！"); 
			            form.accounttext.focus();
			            event.returnValue = false;
			            return;
		            }
		            if (form.accounttext.value.length < 6){
			            alert("您所填寫的帳號少於6碼！"); 
			            form.accounttext.focus();
			            event.returnValue = false;
			            return;
		            }
		            var i;
		            var ch;
		            for (i=0;i< form.accounttext.value.length;i++){
			            ch = form.accounttext.value.charAt(i);
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
				            form.accounttext.focus();
				            event.returnValue = false;
				            return;
		    	        }
		            }
	            }	            
	            if(form.realnametext.value==""){
		            alert("您忘了填寫真實姓名了！"); 
		            form.realnametext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.password1text.value==""){
		            alert("您忘了填寫密碼了！"); 
		            form.password1text.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.password1text.value!=""){
            		
		            if (form.password1text.value.length > 16){
			            alert("您所填寫的密碼超過16碼！"); 
			            form.password1text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else if(form.password1text.value.length < 6){
			            alert("您所填寫的密碼少於6碼！"); 
			            form.password1text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else if(form.password2text.value==""){
			            alert("您忘了填寫密碼確認了！"); 
			            form.password2text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else if(form.password2text.value != form.password1text.value){
			            alert("密碼與密碼確認不符！");
			            form.password2text.focus(); 
			            event.returnValue = false;
			            return;
		            }
		            else{
			            var i;
			            var ch;
			            var digits = "0123456789";
			            var digitflag = false;
			            var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
			            var charflag = false;
			            for (i=0;i< form.password1text.value.length;i++){		
				            ch = form.password1text.value.charAt(i);		
				            if(ch == '\"' || ch == ' ' || ch == '\'' || ch== '\&'){
					            alert("密碼請勿包含『\"』、『'』、『&』或空白"); 
					            form.password1text.focus();
					            event.returnValue = false;
					            return;
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
					            form.password1text.focus();
					            event.returnValue = false;
					            return;
			            }
			            if( !charflag ) {
					            alert("密碼請至少包含一英文字"); 
					            form.password1text.focus();
					            event.returnValue = false;
					            return;
			            }
		            }		
	            }	     
	            if(form.idntext.value==""){
		            alert("您忘了填寫身份證字號了！"); 
		            form.idntext.focus(); 
		            event.returnValue = false;
		            return;
	            }       
	            if(form.member_orgtext.value==""){
		            alert("您忘了填寫所屬機關名稱了！"); 
		            form.member_orgtext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.com_teltext.value==""){
		            alert("您忘了填寫所屬機關電話了！"); 
		            form.com_teltext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.ptitletext.value==""){
		            alert("您忘了填寫職稱了！"); 
		            form.ptitletext.focus(); 
		            event.returnValue = false;
		            return;
	            }	            
	            if(form.emailtext.value==""){
		            alert("您忘了填寫電子郵件了！"); 
		            form.emailtext.focus(); 
		            event.returnValue = false;
		            return;
	            }
	            if(form.emailtext.value.indexOf("@")<=-1){
		            alert("您所填寫的電子郵件格式有誤！"); 
		            form.emailtext.focus(); 
		            event.returnValue = false;
		            return;
	            }
            }              
            </script>

        </div>
    </form>
</body>
</html>
