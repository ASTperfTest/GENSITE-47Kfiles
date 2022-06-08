
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<?xml version="1.0"  encoding="utf-8" ?>
<html xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:hyweb="urn:gip-hyweb-com" xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>農業知識入口網 －小知識串成的大力量－/</title>
</head>

<body>

<h2>搜尋服務</h2>
<form name="SearchForm" method="post" taget="_self" action="">
<label for="search">search</label>
<input type="hidden" name="debug" value="true" />
<input name="Keyword" type="text" accesskey="s" class="txt">
<label for="站內單元" class="ckb"><input name="FromSiteUnit" value="1" type="checkbox" class="ckb" checked>站內單元</label>
<label for="知識庫" class="ckb"><input name="FromKnowledgeTank" value="1" type="checkbox" class="ckb" checked>知識庫</label>
<label for="知識家" class="ckb"><input name="FromKnowledgeHome" value="1" type="checkbox" class="ckb" checked>知識家</label>
<label for="主題館" class="ckb"><input name="FromTopic" value="1" type="checkbox" class="ckb" checked>主題館</label>
<input name="search" type="image" alt="search" class="btn" onClick="javascript:checkSearchForm(0)" src="/xslgip/style3/images/searchBtn.gif">
<input name="search" type="image" alt="search" class="btn" onClick="javascript:checkSearchForm(1)" src="/xslgip/style3/images/SearchBtn2.gif">
</form>
<script language="javascript">
function checkSearchForm(value)
{				
	if( value == 0 ) {
		if( document.SearchForm.Keyword.value == "" ) {
			alert('請輸入查詢值');
			event.returnValue = false;
		}
		else {
			document.SearchForm.action = "/kp.asp?xdURL=Search/SearchResultList.asp";
			document.SearchForm.submit();
		}
	}
	else {
		document.SearchForm.action = "/kp.asp?xdURL=Search/AdvancedSearch.asp";
		document.SearchForm.submit();
	}
}
</script>
</body>
</html>
