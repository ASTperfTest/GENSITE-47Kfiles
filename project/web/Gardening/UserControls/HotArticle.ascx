<%@ Control Language="C#" AutoEventWireup="true" CodeFile="HotArticle.ascx.cs" Inherits="UserControls_HotArticle" %>
	

<table class="title" cellpadding="0" cellspacing="0">
            <tr>
                <td style="width:185px;">
					<img src="img/icon_butterflay.png" />
					<span><asp:Label runat="server" Text="熱門" CssClass="TitleStyle2"></asp:Label> </span>  
                    <span><asp:Label runat="server" Text="討論議題" ></asp:Label> </span>                 
                </td>
                <td  class="more">
                    <span ><asp:HyperLink ID="hlkMore" runat="server" NavigateUrl="~/knowledge/knowledge_garden.aspx?CategoryId=" target="_blank"><img src="img/more.jpg"/></asp:HyperLink><span>  				
                </td>
            </tr>
			<tr style="heigth:10px;">
				<td class="line01">&nbsp;</td>
				<td class="line02">&nbsp;</td>
			</tr>
</table>
	
<div id="table_article">
	<table cellpadding="0" cellspacing="0">
   	  <tr>
          <td><img src="img/corner_left_top.jpg" /></td>
          <td class="border_top"></td>
          <td><img src="img/corner_right_top.jpg" /></td>
      </tr>
      <tr>
        <td class="border_left"></td>
          <td>	
	
	<table width="100%" class="hot_Article">
		<tr>
			<td width="100%" colSpan="2">
				<script language="javascript">
				var articleNumber = <%=articleNumber%>;
				</script>				
			</td>
		</tr>
        <tr>
            <td align="left" width="18%">
				<%=articleType%>
            </td>
			<td align="left" width="67%">
				<%=article%>
				</td>
			<td align="left" width="15%" rowspan="2">
            </td>
        </tr>
    </table>

	
			</td>
          <td class="border_right"></td>
      </tr>
       <tr>
          <td><img src="img/corner_left_button.jpg" /></td>
          <td class="border_button"></td>
          <td><img src="img/corner_right_button.jpg" /></td>
      </tr>
    </table>
</div>