<%@ Page Language="VB" MasterPageFile="~/3RSideMasterPage.master" AutoEventWireup="false" CodeFile="CatTreeContent.aspx.vb" Inherits="CatTreeContent" title="Untitled Page" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    
    <td class="kleft">
    <span class="kleftStyle">
		<asp:treeview id="CatTreeView" DataSourceID="CatTreeViewDS" runat="server" NodeIndent="5" ImageSet="BulletedList4" ShowExpandCollapse="False" >
		    <ParentNodeStyle Font-Bold="False" />                
            <SelectedNodeStyle Font-Underline="True" ForeColor="#5555DD" HorizontalPadding="5" VerticalPadding="2" />
            <NodeStyle Font-Names="Verdana" Font-Size="8pt" ForeColor="Black" HorizontalPadding="5" NodeSpacing="1" VerticalPadding="2" />
		    <DataBindings>
		        <asp:TreeNodeBinding DataMember="node" TextField="name" NavigateUrlField="url"/>			        
		    </DataBindings>
		</asp:treeview>
        <asp:XmlDataSource ID="CatTreeViewDS" EnableCaching="false" runat="server"/>
  </span>
	</td>
	<td class="center">
	    <a title="網頁內容資料" class="Accesskey" accesskey="C" href="#" xmlns:hyweb="urn:gip-hyweb-com">:::</a>
	    <asp:Label runat="server" ID="TabText" Text=""></asp:Label>
	    <div class="path" xmlns:hyweb="urn:gip-hyweb-com">目前位置：<a title="首頁" href="/mp.asp?mp=1">首頁</a>&gt;
		<asp:Label runat="server" ID="NavUrlText" Text=""></asp:Label></div>
           
        <h3><asp:Label runat="server" ID="NavTitleText" Text=""></asp:Label></h3>
            
	    <div id="Magazine">
		    <div class="Event">			    
                <asp:Label ID="TableText" runat="server" Text="" />    
			</div>
		</div>
		<div class="" >
			<%=relationsArticle%>
		</div>
	</td>
    <link rel="stylesheet" href="../knowledge/style/pedia.css" type="text/css" />
	<script type="text/javascript"> 	
	  function trim(stringToTrim){ return stringToTrim.replace(/^\s+|\s+$/g,"");}
		function getSelectedText(path) {  
			var alertStr = "";
			if (window.getSelection) {         
				// This technique is the most likely to be standardized.         
				// getSelection() returns a Selection object, which we do not document.         
				alertStr = window.getSelection().toString();
				//textarea的處理
				if( alertStr == '' ){
					alertStr = getTextAreaSelection();
				}  				
			}          
			else if (document.getSelection) {         
				// This is an older, simpler technique that returns a string         
				alertStr = document.getSelection();     
			}     
			else if (document.selection) {         
				// This is the IE-specific technique.         
				// We do not document the IE selection property or TextRange objects.         
				alertStr = document.selection.createRange().text;     
			}
			if ( alertStr.length > 10 ) {
				alert("詞彙長度限制10字以內");
			}
			//else if ( alertStr == "" ) {
			//	alert("請選擇推薦詞彙");
			//}
			else {
			  alertStr = trim(alertStr);
				window.open(encodeURI("/CommendWord/CommendWordAdd.aspx?type=1&word=" + alertStr + "&" + path ),'建議小百科詞彙','resizable=yes,width=565,height=360');
			}			
		} 

		function getTextAreaSelection(){
			var alertStr = '';
			var elementObj = document.getElementsByTagName("textarea");
			var all_length = elementObj.length;      
			for(var i=0 ; i<all_length ; i++){
				if (elementObj[i].selectionStart != undefined && elementObj[i].selectionEnd != undefined) {         
          var start = elementObj[i].selectionStart;         
          var end = elementObj[i].selectionEnd;         
          alertStr = elementObj[i].value.substring(start, end) ;
          elementObj[i].selectionStart = start;
          elementObj[i].selectionEnd = end;
          //將focus指向該element
          elementObj[i].focus();             
				}     
				else alertStr = ''; // Not supported on this browser                                      
			}
			return alertStr ;
		}
  </script>
  <td><%=innerTag%></td>
</asp:Content>

