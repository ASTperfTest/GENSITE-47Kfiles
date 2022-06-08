<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CommendWordAdd.aspx.vb" Inherits="CommendWord_CommendWordAdd" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>農業知識入口網 －小知識串成的大力量－</title>
    <link href="/xslGip/style3/css/4seasons.css" rel="stylesheet" type="text/css" />
</head>

<body>
  <div class="pantology">
    <div class="head"></div>
    <div class="body">
      <h2>推薦小百科詞彙</h2>
		  <form id="form1" runat="server">
    	  <table summary="排版表格" class="type01">
        <tr>
          <th rowspan="2" nowrap="nowrap" scope="row">推薦詞彙：</th>
          <td>
            <label>              
              <asp:TextBox ID="WordTitle" CssClass="txt" runat="server" ></asp:TextBox>
              <br />
              詞彙長度限制10字以內
            </label>
          </td>
        </tr>
        <tr>
          <td>
			      <img src="/xslGip/style3/images/refresh2.gif" alt="檢查重複" width="15" height="14" align="absmiddle" />
            <asp:LinkButton ID="CheckDuplicateLink" OnClientClick="CheckWord()" runat="server">檢查此詞彙是否已被推薦</asp:LinkButton>(已被推薦之詞彙無法重覆推薦)
			    </td>
        </tr>
        <tr>
          <th nowrap="nowrap" scope="row">詞彙出處：</th>
          <td><asp:Label ID="PathText" runat="server" Text=""></asp:Label></td>
        </tr>
        <tr>
          <th rowspan="2" nowrap="nowrap" scope="row">機器人辨識：</th>
          <td><img src="/CaptchaImage/JpegImage.aspx" alt="機器人辨識" width="200" height="50"></td>
        </tr>
        <tr>
          <td>
            <asp:TextBox ID="CodeNumberTextBox" runat="server" />
            <asp:Label ID="errorshow" Visible="false" ForeColor="red" runat="server" />            
            <br />
            請輸入上方圖片中的數字
          </td>
        </tr>
        </table>
        <!--按鈕 start-->
		    <div class="s01">
          <asp:Button ID="SubmitBtn" runat="server" Text="送出" CssClass="button01" />		                
        </div>
        <!--按鈕 end-->
		  </form>
      <!--填寫欄位form結束-->
    </div>
    <div class="foot"></div>
  </div>
  <script language="javascript" type="text/javascript" >
  
  function CheckWord() {
        
    var data = document.getElementById("WordTitle").value ;
        
    data = data.replace(/(^\s*)|(\s*$)/g,"");
               
    if ( data != "" ) { 
      CallServer(data);  
    }
    else {
      alert("不允許空白");
      return false;
    }
  }
  function ReceiveServerData(rValue) {
    // Y:重複
    if ( rValue=="Y" ) { 
      alert("推薦的詞彙重複");
      return false; 
    }
    else { 
      alert("推薦的詞彙沒有重複");       
      return true; 
    }               
  }
  </script>
</body>
</html>
