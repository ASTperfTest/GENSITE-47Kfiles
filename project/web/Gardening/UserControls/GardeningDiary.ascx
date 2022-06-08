<%@ Control Language="C#" AutoEventWireup="true" CodeFile="GardeningDiary.ascx.cs" Inherits="GardeningDiary" %>
<link href="~/Gardening/css/index.css" rel="stylesheet" type="text/css" />

<table class="title" cellpadding="0" cellspacing="0">
            <tr>
                <td style=" width:140px;">
					<img src="img/icon_cup.png" />
                    <span><asp:Label ID="lbHeaderTitle" runat="server" Text="最新日誌" ></asp:Label> </span>                   
                </td>
                <td  class="more">
                    <span ><asp:HyperLink ID="hlkMore" runat="server" Visible="false" NavigateUrl="http://kminter.coa.gov.tw/gardening/sysentrylist.aspx"><img src="img/more.jpg"/></asp:HyperLink><span>  				
                </td>
            </tr>
			<tr style="heigth:10px;">
				<td class="line01">&nbsp;</td>
				<td class="line02">&nbsp;</td>
			</tr>
</table>
<div id="table_diary">
	<table cellpadding="0" cellspacing="0">
   	  <tr>
          <td><img src="img/corner2_left_top.jpg" /></td>
          <td class="border2_top"></td>
          <td><img src="img/corner2_right_top.jpg" /></td>
      </tr>
      <tr>
        <td class="border2_left"></td>
          <td>	 
<asp:DataList ID="DiaryDataList" runat="server" RepeatColumns="1" >
    <ItemTemplate>

			<asp:Image ID="DiaryIcon" runat="server"  ImageUrl="~/Gardening/images/diary-icon.png" />
			<asp:HyperLink ID="DiaryTitle" runat="server" NavigateUrl='<%# Eval("ImageUri") %>' ><%# Eval("TITLE") %></asp:HyperLink>
			<asp:Label ID="DiaryDate" runat="server" Text='<%# Eval("LastModifyDateTime") %>'></asp:Label>

				
    </ItemTemplate>   
   <HeaderTemplate>       
    </HeaderTemplate>
</asp:DataList>
		
		</td>
          <td class="border2_right"></td>
      </tr>
       <tr>
          <td><img src="img/corner2_left_button.jpg" /></td>
          <td class="border2_button"></td>
          <td><img src="img/corner2_right_button.jpg" /></td>
      </tr>
    </table>
</div>