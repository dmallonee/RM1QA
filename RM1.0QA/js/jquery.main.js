$(function(){
	jQuery('form').customForm({
		disabled: 'disabled'
	});
	initOpen({
		wrap: '.info-box .accaunt-holder ',
		link: 'a.name',	
		box: 'div.list-accaunt-hold',
		openClass: 'active'
	});
	initOpenBox();
	initOpen({
		wrap: '.menu-top >.has-drop',
		link: '>a',
		box: 'ul.drop-list',
		openClass: 'open-box'
	});
	/*$('.area').customScrollV({
		lineWidth: 7
	});*/
	initTable();
	areaBox();
});

function initOpenBox(){
	$('.holder-search').each(function(){
		var hold = $(this);
		var link = hold.find('a.show-more');
		var box = hold.find('div.box-search');
		if(!hold.hasClass('active')){
			box.css({display: 'none'});
		}
		link.click(function(){
			if(!hold.hasClass('active')){
				hold.addClass('active');
				box.slideDown(300);
			}
			else{
				box.slideUp(300, function(){
					hold.removeClass('active');
				});
			}
			return false;
		});
	});
}
function areaBox(){
	var shift = false, ctrl = false;
	$('.area-box').each(function(){
		var hold = $(this);
		var select = hold.find('.area');
		var add = hold.find('.right-side > a');
		var addAll = hold.find('.right-all > a');
		var rem = hold.find('.left-side > a');
		var remAll = hold.find('.left-all > a');
		
		add.click(function(){
			select.eq(-1).find('select').append(select.eq(0).find('select > option').filter(':selected'));
			createSelect();
			return false;
		});
		addAll.click(function(){
			select.eq(-1).find('select').append(select.eq(0).find('select > option'));
			createSelect();
			return false;
		});
		rem.click(function(){
			select.eq(0).find('select').append(select.eq(-1).find('select > option').filter(':selected'));
			createSelect();
			return false;
		});
		remAll.click(function(){
			select.eq(0).find('select').append(select.eq(-1).find('select > option'));
			createSelect();
			return false;
		});
		
		function createSelect(){
			select.each(function(){
				var wrap = $(this);
				var sel = wrap.find('select');
				var opt = wrap.find('option');
				var text = '<ul class="list-items">';
				var li;
				var active = 9999;
				
				sel.css({position: 'absolute', left:-9999,top:-9999});
				
				opt.each(function(){
					text += '<li><a href="#">'+$(this).text()+'</a></li>';
				});
				text = $(text + '</ul>');
				
				wrap.find('ul.list-items').remove();
				sel.after(text);
				
				li = text.find('li');
				
				li.unbind('click').click(function(){
					li.removeClass('active');
					$(this).addClass('active');
					if(ctrl){
						active = li.index(this);
						if (!$(this).hasClass('selected')) {
							$(this).addClass('selected');
							opt.eq(active).prop('selected', true);
						}
						else {
							$(this).removeClass('selected');
							opt.eq(active).prop('selected', false);
						}
					}
					else{
						if(shift){
							var g = li.index(this);
							
							if(active == 9999){
								$(this).addClass('selected');
								active = li.index(this);
								opt.eq(active).prop('selected', true);
							}
							else{
								li.removeClass('selected');
								opt.prop('selected', false);
								if(g >=active){
									for (var i = active; i <= g; i++){
										li.eq(i).addClass('selected');
										opt.eq(i).prop('selected', true);
									};
								}
								else{
									for (var i = active; i >= g; i--){
										li.eq(i).addClass('selected');
										opt.eq(i).prop('selected', true);
									};
								}
							}
						}
						else{
							li.removeClass('selected');
							opt.prop('selected', false);
							$(this).addClass('selected');
							active = li.index(this);
							opt.eq(active).prop('selected', true);
						}
					}
					return false;
				});
				wrap.customScrollV({
					lineWidth: 7
				});
			});
			
		}
		createSelect();
	});
				
	$(document).keydown(function(e){
		ctrl = e.ctrlKey;
		shift = e.shiftKey;
	}).keyup(function(e){
		ctrl = e.ctrlKey;
		shift = e.shiftKey;
	});
}

function initTable(){
	$('.table-box table').each(function(){
		var hold = $(this);
		var tr = hold.find('tr:has(td)');
		var box = hold.find('input:checkbox');
		var active = box.filter(':checked');
		active.parents('tr').addClass('active');
		active.parents('tr').prev().addClass('view');
		if(box.prop('checked')){
			tr.eq($(this)).addClass('active');
		}
		tr.removeClass('odd').each(function(i){
			if(i%2) $(this).addClass('odd');
		});
		box.change(function(){
			if($(this).prop('checked')){
				
				tr.eq(box.index(this)).addClass('active');
				tr.eq(box.index(this)).prev().addClass('view');
			}
			else {
				tr.eq(box.index(this)).removeClass('active');
				tr.eq(box.index(this)).prev().removeClass('view');
			}
		});
	});
}
function initOpen(obj){
	$(obj.wrap).each(function(){
		var hold = $(this);
		var link = hold.find(obj.link);
		var box = hold.find(obj.box);
		if(!hold.hasClass(obj.openClass)){
			box.css({display: 'none'});
		}
		link.click(function(){
			if(!hold.hasClass(obj.openClass)){
				hold.addClass(obj.openClass);
				box.slideDown(300);
			}
			else{
				box.slideUp(300, function(){
					hold.removeClass(obj.openClass);
				});
			}
			return false;
		});
		
		$(document).bind('click touchstart mousedown', function(e){
			if(!($(e.target).parents().index(hold) != -1 || $(e.target).index(hold) != -1)){
				box.slideUp(300, function(){
					hold.removeClass(obj.openClass);
				});
			}
		});
	});
}
/**
 * jQuery Vertical Custom Scroll min v1.0.0
 * Copyright (c) 2013 JetCoders
 * email: yuriy.shpak@jetcoders.com
 * www: JetCoders.com
 * Licensed under the MIT License:
 * http://www.opensource.org/licenses/mit-license.php
 **/

jQuery.uaMatch=function(ua){ua=ua.toLowerCase();var match=/(chrome)[ \/]([\w.]+)/.exec(ua)||/(webkit)[ \/]([\w.]+)/.exec(ua)||/(opera)(?:.*version|)[ \/]([\w.]+)/.exec(ua)||/(msie) ([\w.]+)/.exec(ua)||ua.indexOf("compatible")<0&&/(mozilla)(?:.*? rv:([\w.]+)|)/.exec(ua)||[];return{browser:match[1]||"",version:match[2]||"0"};};if(!jQuery.browser){matched=jQuery.uaMatch(navigator.userAgent);browser={};if(matched.browser){browser[matched.browser]=true;browser.version=matched.version;}if(browser.chrome){browser.webkit=true;}else if(browser.webkit){browser.safari=true;}jQuery.browser=browser;};var types=['DOMMouseScroll','mousewheel'];if($.event.fixHooks){for(var i=types.length;i;){$.event.fixHooks[types[--i]]=$.event.mouseHooks;}}$.event.special.mousewheel={setup:function(){if(this.addEventListener){for(var i=types.length;i;){this.addEventListener(types[--i],handler,false);}}else{this.onmousewheel=handler;}},teardown:function(){if(this.removeEventListener){for(var i=types.length;i;){this.removeEventListener(types[--i],handler,false);}}else{this.onmousewheel=null;}}};$.fn.extend({mousewheel:function(fn){return fn?this.bind("mousewheel",fn):this.trigger("mousewheel");},unmousewheel:function(fn){return this.unbind("mousewheel",fn);}});
function handler(event){var orgEvent=event||window.event,args=[].slice.call(arguments,1),delta=0,returnValue=true,deltaX=0,deltaY=0;event=$.event.fix(orgEvent);event.type="mousewheel";if(orgEvent.wheelDelta){delta=orgEvent.wheelDelta/120;}if(orgEvent.detail){delta=-orgEvent.detail/3;}deltaY=delta;if(orgEvent.axis!==undefined&&orgEvent.axis===orgEvent.HORIZONTAL_AXIS){deltaY=0;deltaX=-1*delta;}if(orgEvent.wheelDeltaY!==undefined){deltaY=orgEvent.wheelDeltaY/120;}if(orgEvent.wheelDeltaX!==undefined){deltaX=-1*orgEvent.wheelDeltaX/120;}args.unshift(event,delta,deltaX,deltaY);return($.event.dispatch||$.event.handle).apply(this,args);}jQuery.easing['jswing']=jQuery.easing['swing'];jQuery.extend(jQuery.easing,{def:'easeOutQuad',swing:function(x,t,b,c,d){return jQuery.easing[jQuery.easing.def](x,t,b,c,d);},easeOutQuad:function(x,t,b,c,d){return-c*(t/=d)*(t-2)+b;},easeOutCirc:function(x,t,b,c,d){return c*Math.sqrt(1-(t=t/d-1)*t)+b;}});jQuery.fn.customScrollV=function(_options){var _options=jQuery.extend({lineWidth:16},_options);return this.each(function(){var _box=jQuery(this);if(_box.is(':visible')){if(_box.children('.scroll-content').length==0){var line_w=_options.lineWidth;var scrollBar=jQuery('<div class="vscroll-bar">'+'	<div class="scroll-up"></div>'+'	<div class="scroll-line">'+'		<div class="scroll-slider">'+'			<div class="scroll-slider-c"></div>'+'		</div>'+'	</div>'+'	<div class="scroll-down"></div>'+'</div>');_box.wrapInner('<div class="scroll-content"><div class="scroll-hold"></div></div>').append(scrollBar);var scrollContent=_box.children('.scroll-content');var scrollSlider=scrollBar.find('.scroll-slider');var scrollSliderH=scrollSlider.parent();var scrollUp=scrollBar.find('.scroll-up');var scrollDown=scrollBar.find('.scroll-down');var box_h=_box.height();var slider_h=0;var slider_f=0;var cont_h=scrollContent.height();var _f=false;var _f1=false;var _f2=true;var _t1,_t2,_s1,_s2;var kkk=0,start=0,_time,flag=true;_box.css({position:'relative',overflow:'hidden',height:box_h});scrollContent.css({position:'absolute',top:0,left:0,zIndex:1,height:'auto'});scrollBar.css({position:'absolute',top:0,right:0,zIndex:2,width:line_w,height:box_h,overflow:'hidden'});scrollUp.css({width:line_w,height:line_w,overflow:'hidden',cursor:'pointer'});scrollDown.css({width:line_w,height:line_w,overflow:'hidden',cursor:'pointer'});slider_h=scrollBar.height();if(scrollUp.is(':visible'))slider_h-=scrollUp.height();if(scrollDown.is(':visible'))slider_h-=scrollDown.height();scrollSliderH.css({position:'relative',width:line_w,height:slider_h,overflow:'hidden'});slider_h=0;scrollSlider.css({position:'absolute',top:0,left:0,width:line_w,height:slider_h,overflow:'hidden',cursor:'pointer'});box_h=_box.height();cont_h=scrollContent.height();if(box_h<cont_h){_f=true;slider_h=Math.round(box_h/cont_h*scrollSliderH.height());if(slider_h<5)slider_h=5;scrollSlider.height(slider_h);slider_h=scrollSlider.outerHeight();slider_f=(cont_h-box_h)/(scrollSliderH.height()-slider_h);_s1=(scrollSliderH.height()-slider_h)/15;_s2=(scrollSliderH.height()-slider_h)/3;scrollContent.children('.scroll-hold').css('padding-right',scrollSliderH.width());}else{_f=false;scrollBar.hide();scrollContent.css({width:_box.width(),top:0,left:0});scrollContent.children('.scroll-hold').css('padding-right',0);};var _top=0;scrollUp.bind('mousedown',function(){_top-=_s1;scrollCont();_t1=setTimeout(function(){_t2=setInterval(function(){_top-=4/slider_f;scrollCont();},20);},500);return false;}).mouseup(function(){if(_t1)clearTimeout(_t1);if(_t2)clearInterval(_t2);}).mouseleave(function(){if(_t1)clearTimeout(_t1);if(_t2)clearInterval(_t2);});scrollDown.bind('mousedown',function(){_top+=_s1;scrollCont();_t1=setTimeout(function(){_t2=setInterval(function(){_top+=4/slider_f;scrollCont();},20);},500);return false;}).mouseup(function(){if(_t1)clearTimeout(_t1);if(_t2)clearInterval(_t2);}).mouseleave(function(){if(_t1)clearTimeout(_t1);if(_t2)clearInterval(_t2);});scrollSliderH.click(function(e){if(_f2){_top=e.pageY-scrollSliderH.offset().top-scrollSlider.outerHeight()/2;scrollCont();}else{_f2=true;};});var t_y=0;var tttt_f=(jQuery.browser.msie)?(true):(false);scrollSlider.mousedown(function(e){t_y=e.pageY-jQuery(this).position().top;_f1=true;return false;}).mouseup(function(){_f1=false;});jQuery('body').bind('mousemove',function(e){if(_f1){_f2=false;_top=e.pageY-t_y;if(tttt_f)document.selection.empty();scrollCont();}}).mouseup(function(){_f1=false;});scrollSlider.bind('touchstart',function(e){if(_time)clearTimeout(_time);scrollSlider.stop();scrollContent.stop();kkk=e.originalEvent.pageY;e.preventDefault();e.stopPropagation();return false;}).bind('touchmove',function(e){if(_f){_f=false;if(kkk>e.originalEvent.pageY)_top-=1*Math.abs(e.originalEvent.pageY-kkk);else _top-=-1*Math.abs(e.originalEvent.pageY-kkk);scrollCont();kkk=e.originalEvent.pageY;_f=true;if((_top>0)&&(_top+slider_h<scrollSliderH.height())){return false;}}e.preventDefault();e.stopPropagation();return false;}).bind('touchend',function(e){e.preventDefault();e.stopPropagation();return false;});_box.bind('touchstart',function(e){if(_time)clearTimeout(_time);scrollSlider.stop();scrollContent.stop();kkk=e.originalEvent.pageY;start=kkk;flag=true;}).bind('touchend',function(e){if(flag&&Math.abs(start-kkk)>80){_top+=(start-kkk)/3;if(_top<0)_top=0;else if(_top+slider_h>scrollSliderH.height())_top=scrollSliderH.height()-slider_h;scrollSlider.animate({top:_top},{queue:false,easing:'easeOutCirc',duration:300*Math.abs(start-kkk)/40});scrollContent.animate({top:-_top*slider_f},{queue:false,easing:'easeOutCirc',duration:300*Math.abs(start-kkk)/40});}e.preventDefault();e.stopPropagation();return false;}).bind('touchmove',function(e){if(_f){_f=false;if(kkk>e.originalEvent.pageY)_top-=-1*Math.abs(e.originalEvent.pageY-kkk)/(cont_h/box_h);else _top-=1*Math.abs(e.originalEvent.pageY-kkk)/(cont_h/box_h);scrollCont();kkk=e.originalEvent.pageY;_f=true;_time=setTimeout(function(){flag=false;},200);if((_top>0)&&(_top+slider_h<scrollSliderH.height())){if(start+30<start+Math.abs(e.originalEvent.pageY))return false;}}if(start+30<start+Math.abs(e.originalEvent.pageY)){e.preventDefault();e.stopPropagation();return false;}});scrollUp.bind('touchstart',function(){_top-=_s1;scrollCont();e.preventDefault();e.stopPropagation();return false;}).bind('touchend',function(e){e.preventDefault();e.stopPropagation();return false;}).bind('touchmove',function(e){e.preventDefault();e.stopPropagation();return false;});scrollDown.bind('touchstart',function(){_top+=_s1;scrollCont();e.preventDefault();e.stopPropagation();return false;}).bind('touchend',function(e){e.preventDefault();e.stopPropagation();return false;}).bind('touchmove',function(e){e.preventDefault();e.stopPropagation();return false;});document.body.onselectstart=function(){if(_f1)return false;};if(!_box.hasClass('not-scroll')){_box.bind('mousewheel',function(event,delta){if(_f){_top-=delta*_s1;scrollCont();if((_top>0)&&(_top+slider_h<scrollSliderH.height()))return false;}});};function scrollCont(){if(_top<0)_top=0;else if(_top+slider_h>scrollSliderH.height())_top=scrollSliderH.height()-slider_h;scrollSlider.css('top',_top);scrollContent.css('top',-_top*slider_f);};this.scrollResize=function(){box_h=_box.height();cont_h=scrollContent.height();if(box_h<cont_h){_f=true;scrollBar.show();scrollBar.height(box_h);slider_h=scrollBar.height();if(scrollUp.is(':visible'))slider_h-=scrollUp.height();if(scrollDown.is(':visible'))slider_h-=scrollDown.height();scrollSliderH.height(slider_h);slider_h=Math.round(box_h/cont_h*scrollSliderH.height());if(slider_h<5)slider_h=5;scrollSlider.height(slider_h);slider_h=scrollSlider.outerHeight();slider_f=(cont_h-box_h)/(scrollSliderH.height()-slider_h);if(cont_h+scrollContent.position().top<box_h)scrollContent.css('top',-(cont_h-box_h));_top=-scrollContent.position().top/slider_f;scrollSlider.css('top',_top);_s1=(scrollSliderH.height()-slider_h)/15;_s2=(scrollSliderH.height()-slider_h)/3;scrollContent.children('.scroll-hold').css('padding-right',scrollSliderH.width());}else{_f=false;scrollBar.hide();scrollContent.css({top:0,left:0});scrollContent.children('.scroll-hold').css('padding-right',0);};};setInterval(function(){if(_box.is(':visible')&&cont_h!=scrollContent.height())_box.get(0).scrollResize();},200);}else{this.scrollResize();};};})};


/**
 * jQuery Custom Form min v1.0.3
 * Copyright (c) 2012 JetCoders
 * email: yuriy.shpak@jetcoders.com
 * www: JetCoders.com
 * Licensed under the MIT License:
 * http://www.opensource.org/licenses/mit-license.php
 **/

;jQuery.fn.customForm=jQuery.customForm=function(e){function r(e,t,r){e.not(".outtaHere").each(function(){var e=$(this);var t=jQuery(n.select.structure);var i=t.find(n.select.text);var s=t.find(n.select.btn);var o=t.find("."+n.disabled).hide();var u=jQuery(n.select.optStructure);var a=u.find(n.select.optList);var f="";var l;if(e.prop("disabled"))o.show();e.find("option").each(function(){var t=jQuery(this);if(t.val()==e.val()){i.html(t.html());f+='<li data-value="'+t.val()+'" class="selected"><a href="#">'+t.html()+"</a></li>"}else f+='<li data-value="'+t.val()+'"><a href="#">'+t.html()+"</a></li>"});a.append(f).find("a").click(function(){a.find("li").removeClass("selected");jQuery(this).parent().addClass("selected");e.val(jQuery(this).parent().attr("data-value"));i.html(jQuery(this).html());e.change();t.removeClass(n.hoverClass);u.hide();return false});t.width(e.outerWidth());t.insertBefore(e);t.addClass(e.attr("class"));u.css({width:e.outerWidth(),display:"none",position:"absolute"});u.addClass(e.attr("class"));jQuery(document.body).append(u);t.hover(function(){if(l)clearTimeout(l)},function(){l=setTimeout(function(){t.removeClass(n.hoverClass);u.hide()},200)});u.hover(function(){if(l)clearTimeout(l)},function(){l=setTimeout(function(){t.removeClass(n.hoverClass);u.hide()},200)});s.click(function(){if(u.is(":visible")){t.removeClass(n.hoverClass);u.hide()}else{t.addClass(n.hoverClass);u.children("ul").css({height:"auto",overflow:"hidden"});e.removeClass("outtaHere");u.css({width:e.outerWidth()});e.addClass("outtaHere");u.css({top:t.offset().top+t.outerHeight(),left:t.offset().left,display:"block"});t.focus();if(n.select.maxHeight&&u.children("ul").height()>n.select.maxHeight)u.children("ul").css({height:n.select.maxHeight,overflow:"auto"})}return false});r.click(function(){setTimeout(function(){e.find("option").each(function(t){var n=jQuery(this);if(n.val()==e.val()){i.html(n.html());a.find("li").removeClass("selected");a.find("li").eq(t).addClass("selected")}})},10)});e.change(function(){if(u.is(":hidden")){e.find("option").each(function(t){var n=jQuery(this);if(n.val()==e.val()){i.html(n.html());a.find("li").removeClass("selected");a.find("li").eq(t).addClass("selected")}})}});$(window).resize(function(){e.removeClass("outtaHere");t.width(Math.floor(e.outerWidth()));e.addClass("outtaHere")})}).addClass("outtaHere")}function i(e,t,r){e.each(function(){var e=$(this);this._label=$("label[for="+e.attr("id")+"]").length?$("label[for="+e.attr("id")+"]"):e.parents("label");if(!e.hasClass("outtaHere")&&e.is(":radio")){var t=jQuery(n.radio.structure);t.addClass(e.attr("class"));this._replaced=t;if(e.is(":disabled")){t.addClass(n.disabled);if(e.is(":checked"))t.addClass("disabledChecked")}else if(e.is(":checked")){t.addClass(n.radio.checked);this._label.addClass("checked")}else{t.addClass(n.radio.defaultArea);this._label.removeClass("checked")}t.click(function(){if(jQuery(this).hasClass(n.radio.defaultArea)){e.change();e.prop("checked",true);s(e.get(0))}});r.click(function(){setTimeout(function(){if(e.is(":checked"))t.removeClass(n.radio.defaultArea+" "+n.radio.checked).addClass(n.radio.checked);else t.removeClass(n.radio.defaultArea+" "+n.radio.checked).addClass(n.radio.defaultArea)},10)});e.click(function(){s(this)});t.insertBefore(e);e.addClass("outtaHere")}})}function s(e){jQuery('input:radio[name="'+jQuery(e).attr("name")+'"]').not(e).each(function(){if(this._replaced&&!jQuery(this).is(":disabled")){this._replaced.removeClass(n.radio.defaultArea+" "+n.radio.checked).addClass(n.radio.defaultArea);this._label.removeClass("checked")}});e._replaced.removeClass(n.radio.defaultArea+" "+n.radio.checked).addClass(n.radio.checked);e._label.addClass("checked");jQuery(e).trigger("change")}function o(e,t,r){e.each(function(){var e=$(this);this._label=$("label[for="+e.attr("id")+"]").length?$("label[for="+e.attr("id")+"]"):e.parents("label");if(!e.hasClass("outtaHere")&&e.is(":checkbox")){var t=jQuery(n.checkbox.structure);t.addClass(e.attr("class"));this._replaced=t;if(e.is(":disabled")){t.addClass(n.disabled);if(e.is(":checked"))t.addClass("disabledChecked")}else if(e.is(":checked")){t.addClass(n.checkbox.checked);this._label.addClass("checked")}else{t.addClass(n.checkbox.defaultArea);this._label.removeClass("checked")}t.click(function(){if(!t.hasClass("disabled")&&!t.parents("label").length){if(e.is(":checked"))e.prop("checked",false);else e.prop("checked",true);u(e)}});r.click(function(){setTimeout(function(){u(e)},10)});e.click(function(){u(e)});t.insertBefore(e);e.addClass("outtaHere");t.parents("label").click(function(){if(!t.hasClass("disabled")){if(e.is(":checked"))e.prop("checked",false);else e.prop("checked",true);u(e)}return false})}})}function u(e){if(e.is(":checked")){e.get(0)._replaced.removeClass(n.checkbox.defaultArea+" "+n.checkbox.checked).addClass(n.checkbox.checked);e.get(0)._label.addClass("checked")}else{e.get(0)._replaced.removeClass(n.checkbox.defaultArea+" "+n.checkbox.checked).addClass(n.checkbox.defaultArea);e.get(0)._label.removeClass("checked")}e.trigger("change")}var t=this;if(typeof t=="function")t=$(document);var n=jQuery.extend(true,{select:{elements:"select.customSelect",structure:'<div class="selectArea"><a href="#" class="selectButton"><span class="center"></span><span class="right"> </span></a><div class="disabled"></div></div>',text:".center",btn:".selectButton",optStructure:'<div class="selectOptions"><ul></ul></div>',maxHeight:false,optList:"ul"},radio:{elements:"input.customRadio",structure:"<div></div>",defaultArea:"radioArea",checked:"radioAreaChecked"},checkbox:{elements:"input.customCheckbox",structure:"<div></div>",defaultArea:"checkboxArea",checked:"checkboxAreaChecked"},disabled:"disabled",hoverClass:"hover"},e);return t.each(function(){var e=jQuery(this);var t=jQuery();if(this!==document)t=e.find("input:reset, button[type=reset]");r(e.find(n.select.elements),e,t);i(e.find(n.radio.elements),e,t);o(e.find(n.checkbox.elements),e,t)})};
