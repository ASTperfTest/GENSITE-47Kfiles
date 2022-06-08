function Show_Hot_Article(id)
{
	var as = "#hot_article_";
	var at ="#hot_articleTitle_";
    for(var i=0;i< articleNumber;i++){
			var ah = "#hot_article_" + i;
			if (i == id){
				$(as+i).show();
				$(at+i).css("background-image","url(images/btn_green.png)");
			}else{
			$(as+i).hide();
			$(at+i).css("background-image","url(images/btn_gray.png)");
			}
        }
}