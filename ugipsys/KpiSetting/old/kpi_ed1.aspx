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
	    <td class="FormName" align="left" style="width: 397px">�Ʀ�]�޲z&nbsp;<font size=2>�i�s�סj</td>
        <td class="FormLink" align="right" style="width: 374px"><A href="Javascript:window.history.back();" title="�^�e��">�^�e��</A>	    </td>
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
  <th style="width: 397px"><font color="red">*</font>�Ʀ�]�W�١G</th>
               <td class="eTableContent" style="width: 374px">
                   
                   <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox></td>
</tr>

<tr>
 <th style="width: 397px">�O�_���}�G</th>
  <td class="eTableContent" style="width: 374px" >
  <select id="Select2"  runat="server" >
      <option>�����}</option>
      <option selected="selected">���}</option>
                    </select> 
   <%-- <asp:Button ID="Button1"  runat="server" Text="�ק�" class="cbutton" />--%>


</table>
  <table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
 <td align="center">     
        <asp:Button ID="Button1"  runat="server" Text="�s�צs��" class="cbutton" />
        <input type=button value ="�R�@��" class="cbutton" onClick="formDelSubmit()">   
        <input type=button value ="���@��" class="cbutton" onClick="resetForm()">
        <input type=button value ="�^�e��" class="cbutton" onClick="Vbs:history.back">
 </td>
</tr>
</table>

   
        
        
        
        
        
        
        
        
        
        
        
        </asp:Content>