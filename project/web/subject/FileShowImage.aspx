<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FileShowImage.aspx.cs" Inherits="SubjectTest" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--
		Supersized - Fullscreen Slideshow jQuery Plugin
		Version 3.1.3
		www.buildinternet.com/project/supersized
		
		By Sam Dunn / One Mighty Roar (www.onemightyroar.com)
		Released under MIT License / GPL License
	-->
<head>
    <title></title>
    <link rel="stylesheet" href="css/supersized.css" type="text/css" media="screen" />

    <script type="text/javascript" src="../js/jquery-1.6.1.min.js"></script>

    <script type="text/javascript" src="js/supersized.3.1.3.js"></script>

    <script type="text/javascript">
		var isPlaying=1;
	
	
        var mySlide_interval = 3000;

        jQuery(function($) {
            $.supersized({

                //Functionality
                slideshow: 1, 	//Slideshow on/off
                autoplay: 1, 	//Slideshow starts playing automatically
                start_slide: 1, 	//Start slide (0 is random)
                random: 0, 	//Randomize slide order (Ignores start slide)
                slide_interval: mySlide_interval, //Length between transitions
                transition: 1, 		//0-None, 1-Fade, 2-Slide Top, 3-Slide Right, 4-Slide Bottom, 5-Slide Left, 6-Carousel Right, 7-Carousel Left
                transition_speed: 500, //Speed of transition
                new_window: 1, 	//Image links open in new window/tab
                pause_hover: 0, 	//Pause slideshow on hover
                keyboard_nav: 1, 	//Keyboard navigation on/off
                performance: 1, 	//0-Normal, 1-Hybrid speed/quality, 2-Optimizes image quality, 3-Optimizes transition speed // (Only works for Firefox/IE, not Webkit)
                image_protect: 1, 	//Disables image dragging and right click with Javascript
                image_path: 'images/', //Default image path

                //Size & Position
                min_width: 0, 	//Min width allowed (in pixels)
                min_height: 0, 	//Min height allowed (in pixels)
                vertical_center: 1, 	//Vertically center background
                horizontal_center: 1, 	//Horizontally center background
                fit_portrait: 1, 	//Portrait images will not exceed browser height
                fit_landscape: 0, 	//Landscape images will not exceed browser width

                //Components
                navigation: 1, 	//Slideshow controls on/off
                thumbnail_navigation: 1, 	//Thumbnail navigation
                slide_counter: 1, 	//Display slide numbers
                slide_captions: 1, 	//Slide caption (Pull from "title" in slides array)
                slides: <%=result %>
				
				//[		//Slideshow Images
						//{image: '../03231424671[1].jpg', title: '觀賞南瓜', url: '' },
								//{image: '../010181724471.jpg', title: '南瓜植株', url: '' },
								//{image: '../0101817234471.jpg', title: '南瓜田開花', url: '' },
								//{ image: 'http://buildinternet.s3.amazonaws.com/projects/supersized/3.1/slides/quietchaos-kitty.jpg', title: 'Quiet Chaos by Kitty Gallannaugh', url: '' },
								//{ image: 'http://buildinternet.s3.amazonaws.com/projects/supersized/3.1/slides/wanderers-kitty.jpg', title: 'Wanderers by Kitty Gallannaugh', url: '' },
								//{ image: 'http://buildinternet.s3.amazonaws.com/projects/supersized/3.1/slides/apple-kitty.jpg', title: 'Applewood by Kitty Gallannaugh', url: '' }
								//]
            });
        });        

        function Slide_Interval(s) {
            mySlide_interval += s;
            mySlide_interval = Math.max(1000, mySlide_interval);
            mySlide_interval = Math.min(5000, mySlide_interval);
            switch (mySlide_interval) {
                case 1000:
                    $("#speedTitle").text("快");
                    break;
                case 3000:
                    $("#speedTitle").text("中");
                    break;
                case 5000:
                    $("#speedTitle").text("慢");
                    break;
            }
            $('#nextthumb').click();
        }
    </script>



    <style type="text/css">
        /*Demo Styles*/p
        {
            padding: 0 30px 30px 30px;
            color: #fff;
            font: 11pt "Helvetica Neue" , "Helvetica" , Arial, sans-serif;
            text-shadow: #000 0px 1px 0px;
            line-height: 200%;
        }
        p a
        {
            font-size: 10pt;
            text-decoration: none;
            outline: none;
            color: #ddd;
            background: #222;
            border-top: 1px solid #333;
            padding: 5px 8px;
            -moz-border-radius: 3px;
            -webkit-border-radius: 3px;
            border-radius: 3px;
            -moz-box-shadow: 0px 1px 1px #000;
            -webkit-box-shadow: 0px 1px 1px #000;
            box-shadow: 0px 1px 1px #000;
        }
        p a:hover
        {
            background-color: #427cb4;
            border-color: #5c94cb;
            color: #fff;
        }
        h3
        {
            padding: 30px 30px 20px 30px;
        }
        #content
        {
            position: absolute;
            top: 50px;
            left: 50px;
            background: #111;
            background: rgba(0,0,0,0.70);
            width: 360px;
            text-align: left;
        }
        .stamp
        {
            float: right;
            margin: 15px 30px 0 0;
        }
    </style>
</head>
<body>
<div id='dummyspan'>
<embed  Played='1' hidden='true' autostart='true' loop='true'  src='/subject/midi/<%=getmusic %>'>
</div>
    <!--Thumbnail Navigation-->
    <div id="prevthumb">
    </div>
    <div id="nextthumb">
    </div>
    <!--Control Bar-->
    <div id="controls-wrapper">
        <div id="controls">
            <!--Slide counter-->
            <div id="slidecounter">
                <span class="slidenumber"></span>/<span class="totalslides"></span>
            </div>
            <!--Slide captions displayed here-->
            <div id="slidecaption">
            </div>
            <!--Navigation-->
            <div id="navigation">            
            <%if (getmusic != ""){ %>
             <a href="javascript:void(0)" onclick="play(this)"><img src='images/Music_on.png'></a>
            <%}%>
            <img id="pauseplay" src="images/pause_dull.png" width="32" />
              
            </div>   
            <div id="navigationNew">
            <img id="Img4" src="images/back_icon.png" width="20" onclick="Slide_Interval(2000)" alt="調慢" />
                <span>播放速度：</span><span id="speedTitle">中</span>
            <img id="Img5" src="images/next_icon.png" width="20" onclick="Slide_Interval(-2000)" alt="調快" />
  
            </div>               

            <script type="text/javascript">			
                function play(o) {
                    if (isPlaying == 1) {						
                        o.innerHTML = "<img src='images/Music_off.png'>";
                        isPlaying = 0;
                        SoundOff();
                    }
					else{				
                        o.innerHTML = "<img src='images/Music_on.png'>";
                        isPlaying = 1;
                        DHTMLSound();						
                    }                     
                }


            </script>

        </div>
        <!--Logo in bar-->
        </div>
    </div>
	
<script language="JavaScript">
function DHTMLSound(surl) {
document.getElementById("dummyspan").innerHTML=
"<embed  Played='1' hidden='true' autostart='true' loop='true'  src='/subject/midi/<%=getmusic %>'>";
}



function SoundOff(surl) {
document.getElementById("dummyspan").innerHTML=
"<embed  Played='0' hidden='true' autostart='false' loop='false'  src=''>";
}
</script>

	
</body>
</html>
