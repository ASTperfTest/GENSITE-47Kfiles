        var sizeLimit=200;   
   function checkFile() {  //sizeLimit單位:byte

           var f = document.form1;
           var re = /(\.jpg|\.gif)$/i; //允許的圖片副檔名   
       
  
         
            
           if (!re.test(f.FileUpload1.value)) {

                     alert("只允許上傳JPG或GIF影像檔");

           } else {
                     var img = new Image();   
                      
                    
                        
                      img.src = f.FileUpload1.value;   

                  img.onload = showImageDimensions
           }         

} 

function showImageDimensions() { 
     
            
               
           if (this.fileSize > sizeLimit) { 
                     alert("over");
                     alert('您所選擇的檔案大小為 '+ (this.fileSize/1000) +' kb，\n超過了上傳上限 ' + (sizeLimit/1000) + ' kb！\n不允許上傳！'); 

                     document.getElementById("FileUpload1").outerHTML = '<input type="file" name="file1" size="20" id="file1">';  

           } else { 
                     alert("ok");
                     document.MM_returnValue = true;

           }         

} 