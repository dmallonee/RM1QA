
// Original:  Nannette Thacker (nannette@shiningstar.net) -->
// Web Site:  http://www.shiningstar.net -->

//This script and many more are available free online at -->
// The JavaScript Source!! http://javascript.internet.com -->

//
var version4 = (navigator.appVersion.charAt(0) == "4"); 
var popupHandle;
function closePopup() {
if(popupHandle != null && !popupHandle.closed) popupHandle.close();
}
function displayPopup(position,url,name,height,width,evnt) {
// position=1 POPUP: makes screen display up and/or left, down and/or right 
// depending on where cursor falls and size of window to open
// position=2 CENTER: makes screen fall in center
var properties = "toolbar = 0, location = 0, height = " + height;
properties = properties + ", width=" + width;
var leftprop, topprop, screenX, screenY, cursorX, cursorY, padAmt;
if(navigator.appName == "Microsoft Internet Explorer") {
screenY = document.body.offsetHeight;
screenX = window.screen.availWidth;
}
else {
screenY = window.outerHeight
screenX = window.outerWidth
}
if(position == 1)	{ // if POPUP not CENTER
cursorX = evnt.screenX;
cursorY = evnt.screenY;
padAmtX = 10;
padAmtY = 10;
if((cursorY + height + padAmtY) > screenY) {
// make sizes a negative number to move left/up
padAmtY = (-30) + (height * -1);
// if up or to left, make 30 as padding amount
}
if((cursorX + width + padAmtX) > screenX)	{
padAmtX = (-30) + (width * -1);	
// if up or to left, make 30 as padding amount
}
if(navigator.appName == "Microsoft Internet Explorer") {
leftprop = cursorX + padAmtX;
topprop = cursorY + padAmtY;
}
else {
leftprop = (cursorX - pageXOffset + padAmtX);
topprop = (cursorY - pageYOffset + padAmtY);
   }
}
else{
leftvar = (screenX - width) / 2;
rightvar = (screenY - height) / 2;
if(navigator.appName == "Microsoft Internet Explorer") {
leftprop = leftvar;
topprop = rightvar;
}
else {
leftprop = (leftvar - pageXOffset);
topprop = (rightvar - pageYOffset);
   }
}
if(evnt != null) {
properties = properties + ", left = " + leftprop;
properties = properties + ", top = " + topprop;
}
closePopup();
popupHandle = open(url,name,properties);
}