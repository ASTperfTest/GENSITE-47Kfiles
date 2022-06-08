<%@ Page Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" CodeFile="kpi_ed1.aspx.vb" Inherits="kpi_ed1" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    &nbsp; &nbsp;&nbsp;





<table border="0" width="100%" cellspacing="0" cellpadding="0">
	  
	  <tr>
	    <td width="100%" colspan="2">
	      
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  

 
	
	  <tr>
	    <td class="FormName" align="left" style="width: 397px">排行榜管理&nbsp;<font size=2>【編修】</td>
        <td class="FormLink" align="right" style="width: 374px"><A href="Javascript:window.history.back();" title="回前頁">回前頁</A>	    </td>
	  </tr>
	  <tr>
	    <td width="100%" colspan="2">
	      <hr noshade size="1" color="#000080">
	    </td>
	  </tr>
	  <tr>
	    <td class="Formtext" colspan="2" height="15"></td>
	  </tr>  
       <tr>
	       

	  

<tr>
  <th style="width: 397px"><font color="red">*</font>排行榜名稱：</th>
               <td class="eTableContent" style="width: 374px">
                   
                   <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox></td>
</tr>

<tr>
 <th style="width: 397px">是否公開：</th>
  <td class="eTableContent" style="width: 374px" >
  <select id="Select2"  runat="server" >
      <option>不公開</option>
      <option selected="selected">公開</option>
                    </select> 
   <%-- <asp:Button ID="Button1"  runat="server" Text="修改" class="cbutton" />--%>


</table>
  <table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
 <td align="center">     
        <asp:Button ID="Button1"  runat="server" Text="編修存檔" class="cbutton" />
        <input type=button value ="刪　除" class="cbutton" onClick="formDelSubmit()">   
        <input type=button value ="重　填" class="cbutton" onClick="resetForm()">
        <input type=button value ="回前頁" class="cbutton" onClick="Vbs:history.back">
 </td>
</tr>
</table>

   
        
        
        
        
        
        
        
        
        
        
        
        </asp:Content>