<%@ Page Language="C#" CodePage="65001"%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <title>search</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
    <script type="text/javascript"">        
		var hu = window.location.search.substring(1);
		var ft = hu.split("=");
        location.href='http://www.google.com.tw/search?as_q=&hl=zh-TW&num=10&btnG=Google+%E6%90%9C%E5%B0%8B&as_epq=&as_oq='+ft[1]+'&as_eq=&lr=&cr=&as_ft=i&as_filetype=&as_qdr=all&as_occt=any&as_dt=i&as_sitesearch=kmweb.coa.gov.tw&as_rights=&safe=images';
		</script>

</body>
</html>
