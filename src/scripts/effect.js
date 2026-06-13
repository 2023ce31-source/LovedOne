$(window).load(function(){
	$('.loading').fadeOut('fast');
	$('.container').fadeIn('fast');
});

$('document').ready(function(){
	var vw;
	
	$(window).resize(function(){
		vw = $(window).width()/2;
		$('#card *').hide();
		$('#b1,#b2,#b3,#b4,#b5,#b6,#b7').stop();
		
		// Mobile-friendly balloon positioning on resize
		if($(window).width() < 768) {
			$('#b11').animate({top:240, left: 10},500);
			$('#b22').animate({top:240, left: vw-200},500);
			$('#b33').animate({top:240, left: vw-100},500);
			$('#b44').animate({top:240, left: vw},500);
			$('#b55').animate({top:240, left: vw+100},500);
			$('#b66').animate({top:240, left: vw+200},500);
			$('#b77').animate({top:240, left: 2*vw-10},500);
		} else {
			$('#b11').animate({top:240, left: vw-350},500);
			$('#b22').animate({top:240, left: vw-250},500);
			$('#b33').animate({top:240, left: vw-150},500);
			$('#b44').animate({top:240, left: vw-50},500);
			$('#b55').animate({top:240, left: vw+50},500);
			$('#b66').animate({top:240, left: vw+150},500);
			$('#b77').animate({top:240, left: vw+250},500);
		}
	});

	$('#initial').click(function(){
		$(this).fadeOut('slow').delay(50).promise().done(function(){
			$('#turn_on').fadeIn('slow');
		});
	});

	$('#play').click(function(){
		var audio = $('.song')[0];
		audio.play();
		$('#bulb_yellow').addClass('bulb-glow-yellow-after');
		$('#bulb_red').addClass('bulb-glow-red-after');
		$('#bulb_blue').addClass('bulb-glow-blue-after');
		$('#bulb_green').addClass('bulb-glow-green-after');
		$('#bulb_pink').addClass('bulb-glow-pink-after');
		$('#bulb_orange').addClass('bulb-glow-orange-after');
		$('body').css('background-color','#FFF');
		$('body').addClass('peach-after');
		$(this).fadeOut('slow').delay(5000).promise().done(function(){
			$('#bannar_coming').fadeIn('slow');
		});
	});

	$('#turn_on').click(function(){
		$('#bulb_yellow').addClass('bulb-glow-yellow');
		$('#bulb_red').addClass('bulb-glow-red');
		$('#bulb_blue').addClass('bulb-glow-blue');
		$('#bulb_green').addClass('bulb-glow-green');
		$('#bulb_pink').addClass('bulb-glow-pink');
		$('#bulb_orange').addClass('bulb-glow-orange');
		$('body').addClass('peach');
		$("#welcome_dance").remove();
		$(this).fadeOut('slow').delay(5000).promise().done(function(){
			$('#play').fadeIn('slow');
		});
	});

	$('#bannar_coming').click(function(){
		$('#banner').css("display", "block");
		$('.bannar').addClass('center');
		$('.bannar').addClass('bannar-come');
		$(this).fadeOut('slow').delay(5000).promise().done(function(){
			$('#balloons_flying').fadeIn('slow');
		});
	});

	function loopOne() {
		var randleft = Math.floor((Math.random() * $(window).width()) + 1);
		var randtop =  Math.floor((Math.random() * $(window).height()) + 1);
		$('#b1').animate({left:randleft,bottom:randtop},8000,function(){
			loopOne();
		});
	}
	
	function loopTwo() {
		var randleft = Math.floor((Math.random() * $(window).width()) + 1);
		var randtop =  Math.floor((Math.random() * $(window).height()) + 1);
		$('#b2').animate({left:randleft,bottom:randtop},8000,function(){
			loopTwo();
		});
	}
	
	function loopThree() {
		var randleft = Math.floor((Math.random() * $(window).width()) + 1);
		var randtop =  Math.floor((Math.random() * $(window).height()) + 1);
		$('#b3').animate({left:randleft,bottom:randtop},8000,function(){
			loopThree();
		});
	}
	
	function loopFour() {
		var randleft = Math.floor((Math.random() * $(window).width()) + 1);
		var randtop =  Math.floor((Math.random() * $(window).height()) + 1);
		$('#b4').animate({left:randleft,bottom:randtop},8000,function(){
			loopFour();
		});
	}
	
	function loopFive() {
		var randleft = Math.floor((Math.random() * $(window).width()) + 1);
		var randtop =  Math.floor((Math.random() * $(window).height()) + 1);
		$('#b5').animate({left:randleft,bottom:randtop},8000,function(){
			loopFive();
		});
	}

	function loopSix() {
		var randleft = Math.floor((Math.random() * $(window).width()) + 1);
		var randtop =  Math.floor((Math.random() * $(window).height()) + 1);
		$('#b6').animate({left:randleft,bottom:randtop},8000,function(){
			loopSix();
		});
	}
	
	function loopSeven() {
		var randleft = Math.floor((Math.random() * $(window).width()) + 1);
		var randtop =  Math.floor((Math.random() * $(window).height()) + 1);
		$('#b7').animate({left:randleft,bottom:randtop},8000,function(){
			loopSeven();
		});
	}

	$('#balloons_flying').click(function(){
		$('.balloon-border').animate({top:-500},8000);
		$('#b1,#b4,#b5,#b7').addClass('balloons-rotate-behaviour-one');
		$('#b2,#b3,#b6').addClass('balloons-rotate-behaviour-two');
		
		loopOne();
		loopTwo();
		loopThree();
		loopFour();
		loopFive();
		loopSix();
		loopSeven();
		
		$(this).fadeOut('slow').delay(5000).promise().done(function(){
			$('#invite_friends').fadeIn('slow');
		});
	});	

	$('#invite_friends').click(function(){
		vw = $(window).width()/2;
		mid = $(window).height()/2;
		
		// Detect if mobile and adjust scaling - INCREASED from 0.6 to 0.9 for better spacing
		var isMobile = $(window).width() < 768;
		var scale = isMobile ? 0.9 : 1;
		
		// Calculate differences with mobile scaling
		mid_diff = (mid/12) * scale;
		vw_diff = (vw/12) * scale;
		
		// Self (center character) - z-index set in CSS
		$('#self').css({"display": "block", "bottom":'0px',"left":vw+'px'});
		$('#self').animate({top:mid-mid_diff*4,left:vw},3000);

		// Friend 1 - top left
		$('#f1').css({"display": "block", "top":500*Math.random()+'px',"left":'0px'});
		$('#f1').animate({top:mid-mid_diff*3,left:vw-vw_diff*3},3000);

		// Friend 2 - top right
		$('#f2').css({"display": "block", "top":500*Math.random()+'px',"left":'0px'});
		$('#f2').animate({top:mid-mid_diff*3,left:vw+vw_diff*3},3000);

		// Friend 3 - mid-top left
		$('#f3').css({"display": "block", "top":500*Math.random()+'px',"left":'0px'});
		$('#f3').animate({top:mid-mid_diff*2,left:vw-vw_diff*4.5},3000);

		// Friend 4 - mid-top right
		$('#f4').css({"display": "block", "top":500*Math.random()+'px',"left":'0px'});
		$('#f4').animate({top:mid-mid_diff*2,left:vw+vw_diff*4.5},3000);

		// Friend 5 - right side
		$('#f5').css({"display": "block", "top":500*Math.random()+'px',"right":'0px'});
		$('#f5').animate({top:mid-mid_diff,left:vw+vw_diff*5.5},3000);

		// Friend 6 - left side
		$('#f6').css({"display": "block", "top":500*Math.random()+'px',"right":'0px'});
		$('#f6').animate({top:mid-mid_diff,left:vw-vw_diff*5.5},3000);

		// Friend 7 - far left
		$('#f7').css({"display": "block", "top":500*Math.random()+'px',"right":'0px'});
		$('#f7').animate({top:mid,left:vw-vw_diff*6.5},3000);

		// Friend 8 - far right
		$('#f8').css({"display": "block", "top":500*Math.random()+'px',"right":'0px'});
		$('#f8').animate({top:mid,left:vw+vw_diff*6.5},3000);

		// Friend 9 - bottom right
		$('#f9').css({"display": "block", "top":500*Math.random()+'px',"left":'0px'});
		$('#f9').animate({top:mid+mid_diff*3,left:vw+vw_diff*2.5},3000);

		// Friend 10 - bottom far left
		$('#f10').css({"display": "block", "top":500*Math.random()+'px',"right":'0px'});
		$('#f10').animate({top:mid+mid_diff,left:vw-vw_diff*7.5},3000);

		// Friend 11 - bottom far right
		$('#f11').css({"display": "block", "top":500*Math.random()+'px',"right":'0px'});
		$('#f11').animate({top:mid+mid_diff,left:vw+vw_diff*7.5},3000);

		// Friend 12 - bottom left
		$('#f12').css({"display": "block", "top":500*Math.random()+'px',"left":'0px'});
		$('#f12').animate({top:mid+mid_diff*3,left:vw-vw_diff*4.5},3000);

		$(this).fadeOut('slow').delay(5000).promise().done(function(){
			$('#cake_fadein').fadeIn('slow');
		});
	});	

	$('#cake_fadein').click(function(){
		$('#cake_pink').css("display", "block");
		$('#candle').css("display", "block");

		$('#cake_pink').fadeIn('slow');
		$(this).fadeOut('slow').delay(3000).promise().done(function(){
			$('#dj_time').fadeIn('slow');
		});
	});
		
	$('#dj_time').click(function(){
		var audio = $('.song')[0];
		audio.pause();

		var audio1 = $('.song1')[0];
		audio1.play();
		
		$('body').addClass('peach-disco');
		$('#f1').addClass('dance-one');
		$('#f3').addClass('dance-two');
		$('#f2').addClass('dance-one');
		$('#f4').addClass('dance-two');
		$('#f5').addClass('dance-one');
		$('#f6').addClass('dance-one');
		$('#f7').addClass('dance-two');
		$('#f8').addClass('dance-two');
		$('#f11').addClass('dance-two');
		$('#f12').addClass('dance-two');
		$('#self').addClass('dance-one');
		$('#f9').addClass('dance-one');
		$('#f10').addClass('dance-one');
		
		$('#bulb_yellow').addClass('bulb-glow-yellow-disco');
		$('#bulb_red').addClass('bulb-glow-red-disco');
		$('#bulb_blue').addClass('bulb-glow-blue-disco');
		$('#bulb_green').addClass('bulb-glow-green-disco');
		$('#bulb_pink').addClass('bulb-glow-pink-disco');
		$('#bulb_orange').addClass('bulb-glow-orange-disco');
		
		vw = $(window).width()/2;
		var isMobile = $(window).width() < 768;

		$('#b1,#b2,#b3,#b4,#b5,#b6,#b7').stop();
		$('#b1').attr('id','b11');
		$('#b2').attr('id','b22');
		$('#b3').attr('id','b33');
		$('#b4').attr('id','b44');
		$('#b5').attr('id','b55');	
		$('#b6').attr('id','b66');
		$('#b7').attr('id','b77');
		
		// Mobile-friendly balloon positioning
		if(isMobile) {
			$('#b11').animate({top:240, left: 5},500);
			$('#b22').animate({top:240, left: vw*0.3},500);
			$('#b33').animate({top:240, left: vw*0.65},500);
			$('#b44').animate({top:240, left: vw},500);
			$('#b55').animate({top:240, left: vw*1.35},500);
			$('#b66').animate({top:240, left: vw*1.7},500);
			$('#b77').animate({top:240, left: 2*vw-5},500);
		} else {
			$('#b11').animate({top:240, left: 0},500);
			$('#b22').animate({top:240, left: vw-vw*2/3},500);
			$('#b33').animate({top:240, left: vw-vw/3},500);
			$('#b44').animate({top:240, left: vw},500);
			$('#b55').animate({top:240, left: vw+vw/3},500);
			$('#b66').animate({top:240, left: vw+vw*2/3},500);
			$('#b77').animate({top:240, left: 2*vw},500);
		}
		
		$('.balloons').css('opacity','0.9');
		$('.balloons h2').fadeIn(3000);
		$(this).fadeOut('slow').delay(3000).promise().done(function(){
			$('#story').fadeIn('slow');
		});
	});
	
	$('#story').click(function(){
		$(this).fadeOut('slow');
		$('#card').css("display", "block");
		$('#card').show();
		$('#card').css('position', 'absolute');
		$('#card').css('z-index', 9999);
		$('.cake').fadeOut('fast').promise().done(function(){
			$('.message').fadeIn('slow');
		});
		$('#card').fadeIn('slow');

		$(this).fadeOut('slow').delay(3000).promise().done(function(){
			$('#picture_time').fadeIn('slow');
		});
	});

	$('#picture_time').click(function(){
		window.location.replace("surprise.html");
	});
});
