<%@ Page Language="C#" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="index" MasterPageFile="Default.master" %>
<%@ Register Src="UserControls/GardeningExpert.ascx" TagName="GardenExpertDataList" TagPrefix="uc4" %> 
<%@ Register Src="UserControls/GardeningDiary.ascx" TagName="GardenDiaryDataList" TagPrefix="uc2" %>  
<%@ Register Src="UserControls/GardeningBestTopics.ascx" TagName="GardenBestTopicsDataList" TagPrefix="uc3" %>  

<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">


<%@ Register Src="UserControls/TagCloud.ascx" TagName="TagCloudList" TagPrefix="uc1" %> 
	
<div id="left_field">
	<div id="div_best_topics">
		<uc3:GardenBestTopicsDataList  runat="server"/>
	</div>
</div>	
	
<div id="right_field">
	<div id="div_search">
		<table cellpadding="0" cellspacing="0" id="table_search">
		  <tr>
			  <td><img src="img/corner_left_top.jpg" /></td>
			  <td colspan="2" class="border_top"></td>
			  <td><img src="img/corner_right_top.jpg" /></td>
		  </tr>
		  <tr id="search_keyword">
			  <td class="border_left"></td>
			  <td colspan="2" >
				<uc1:TagCloudList  runat="server"/>
			  </td>
			  <td class="border_right"></td>
		  </tr>
		   <tr>
			  <td><img src="img/corner_left_button.jpg" /></td>
			  <td colspan="2" class="border_button"></td>
			  <td><img src="img/corner_right_button.jpg" /></td>
		  </tr>
		</table>
	</div>
		
			

	<div id="div_diary">
		<uc2:GardenDiaryDataList  runat="server"/>
	</div>

	<div id="div_article">
		<asp:PlaceHolder ID="hot_PlaceHolder" runat="server"></asp:PlaceHolder>	
	</div>

</div>
	
	
	



</asp:Content>