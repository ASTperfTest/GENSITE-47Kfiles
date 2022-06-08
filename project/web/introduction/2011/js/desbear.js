			
			$().ready(function() {
			     
			    $('td pre code').each(function() {
				        eval($(this).text());
					});
					$('.radimg').cycle({ 
				    fx: 'shuffle', 
				    speed:  2000, 
					timeout:   200,
				    pause:  0
				 });
 
			   
			  
					    
					
								   
			   
			    $('.tip').toggle(function(){
			                $('.tip').not(this).find('.contentp').slideUp(100);
			                $(this).find('.contentp').slideDown('300'); 
			                $(this).find('.tipar').hide();
			                         
			                              }
			    			,function(){$(this).find('.contentp').slideUp('300'); 
			    			            $(this).find('.tipar').show(); 
			                               });
			                               
			     $('.tip').hover(function(){
			                $('.tip').not(this).find('.contentp').slideUp(100);
			                         
			                              }
			    			,function(){$(this).find('.contentp').slideUp('300'); 
			    			            $(this).find('.tipar').show(); 
			                               });                           
			       
				  							});