<%@ Page Language="C#" AutoEventWireup="true" CodeFile="SearchResult.aspx.cs" Inherits="SearchResult" MasterPageFile="Default.master"%>


<asp:Content ID="Content1" ContentPlaceHolderID="cp" runat="Server">

	<table class="layout">
		<tr>
			<td width="40%">
				<script language="javascript">

				    function moreKnowledge() {
				        var keyword = "<%=Keyword%>";
				        if (keyword != "") {
				            document.SearchForm.Keyword.value = keyword;
				            document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp";
				            document.SearchForm.debug.value = false;
				            document.SearchForm.submit();
				        } else {
				            alert("please insert searchworld");
				        }
				    }
						</script>
			</td>
			<td width="60%" id="search_result_padding_top">
				<table class ="searchtable">
				    <tr>
				        <td>
				            <h3>相關作品</h3>
				        </td>
				    </tr>
					<tr>
						<td>
							<%=searchTopicStr%>
						</td>
					</tr>
					<tr>
				        <td>
				            <h3>相關知識問題</h3>
				        </td>
				    </tr>
					<tr>
						<td>
							<%=searchResultStr%>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		
	</table>
</asp:Content>