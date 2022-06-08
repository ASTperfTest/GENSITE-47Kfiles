<%@ Page Language="C#" AutoEventWireup="true" CodeFile="FileShowPage.aspx.cs" Inherits="FileShowPage" %>

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
    <link rel="stylesheet" href="/js/slideshow/css/supersized.css" type="text/css" media="screen" />
    <script type="text/javascript" src="/js/jquery.js"></script>
    <script type="text/javascript" src="/js/slideshow/js/supersized.3.1.3.js"></script>
    <script type="text/javascript">
        var mySlide_interval = 3000;

        jQuery(function ($) {
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
                image_path: 'img/', //Default image path

                //Size & Position
                min_width: 1, 	//Min width allowed (in pixels)
                min_height: 1, 	//Min height allowed (in pixels)
                vertical_center: 1, 	//Vertically center background
                horizontal_center: 1, 	//Horizontally center background
                fit_portrait: 1, 	//Portrait images will not exceed browser height
                fit_landscape: 1, 	//Landscape images will not exceed browser width

                //Components
                navigation: 1, 	//Slideshow controls on/off
                thumbnail_navigation: 1, 	//Thumbnail navigation
                slide_counter: 1, 	//Display slide numbers
                slide_captions: 1, 	//Slide caption (Pull from "title" in slides array)
                slides: <%=result %>
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
    <!--Thumbnail Navigation-->
    <div id="prevthumb">
    </div>
    <div id="controls-wrapper">
        <div id="controls">
            <div id="slidecounter">
                <span class="slidenumber"></span>/<span class="totalslides"></span>
            </div>
            <!--Slide Captions Displayed here-->
            <div id="slidecaption">
            </div>
        </div>
    </div>
    <div id="nextthumb">
    </div>
</body>
</html>
