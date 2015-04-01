<%@ Language=VBScript %>
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180

	On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2
	Dim adoCmd3	
	Dim adoRS3
	Dim adoCmd4
	Dim adoRS4


	Dim adoPrices
	Dim strUserId
	Dim intRuleId
	Dim strAlertDesc
	Dim datBeginDate


	'strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "select * from org"
		
	'Set adoRS1 = adoCmd.Execute


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="index,follow" name="robots" />
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type" />
<link href="pics/homescreen.png" rel="apple-touch-icon" />
<meta content="minimum-scale=1.0, width=device-width, maximum-scale=0.6667, user-scalable=no" name="viewport" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script src="javascript/functions.js" type="text/javascript"></script>
<title>Rate-Monitor | Org List</title>
<meta content="iPod,iPhone,Webkit,iWebkit,Website,Create,mobile,Tutorial,free" name="Keywords" />
<meta content="Create the classical iphone list feeling with these lists. Add images to make bigger and nicer lists." name="description" />
</head>

<body class="list">

<div id="topbar">
	<div id="leftnav">
		<a href="default.htm"><img alt="home" src="images/home.png" /></a></div>
	<div id="rightnav">
		<a href="musiclist.html">Music list</a></div>
	<div id="title">
		List</div>
</div>
<div id="content">
	<ul class="autolist">
		<li class="title">navigation</li>
		<li><a href="default.htm"><span class="name">Go Home</span><span class="arrow"></span></a></li>
		<li><a href="userlist.html"><span class="name">Go to user list</span><span class="arrow"></span></a></li>
		<li class="title">Music to buy</li>
		<li class="withimage">
		<a class="noeffect" href="http://itunes.apple.com/WebObjects/MZStore.woa/wa/viewAlbum?id=130244757">
		<img alt="test" src="pics/californiacation.jpg" width="90" height="90" /><span class="name">Buy 
		Album on iTunes</span><span class="comment">Californiacation</span><span class="arrow"></span></a></li>
		<li class="title">To get at the supermarket:</li>
		<li class="withimage"><a class="noeffect">
		<img alt="test" src="pics/milk.jpg" width="90" height="90" /><span class="name">Buy 
		Milk</span><span class="comment">Check the date</span></a></li>
		<li><a class="noeffect"><span class="name">Eggs</span></a></li>
		<li><a class="noeffect"><span class="name">Bread</span></a></li>
		<li><a class="noeffect"><span class="name">Cheese</span></a></li>
		<li><a class="noeffect"><span class="name">tomatoes</span></a></li>
		<li><a class="noeffect"><span class="name">Salad</span></a></li>
		<li><a class="noeffect"><span class="name">pie</span></a></li>
		<li><a class="noeffect"><span class="name">a T-shirt</span></a></li>
		<li><a class="noeffect"><span class="name">sandwiches</span></a></li>
		<li><a class="noeffect"><span class="name">computer</span></a></li>
		<li><a class="noeffect"><span class="name">ipod touch</span></a></li>
		<li><a class="noeffect"><span class="name">apples</span></a></li>
		<li><a class="noeffect"><span class="name">Windows Vista</span></a></li>
		<li><a class="noeffect"><span class="name">soup</span></a></li>
		<li><a class="noeffect"><span class="name">Almonds</span></a></li>
		<li><a class="noeffect"><span class="name">Black pepper</span></a></li>
		<li><a class="noeffect"><span class="name">Beef, lean organic</span></a></li>
		<li><a class="noeffect"><span class="name">Turkey</span></a></li>
		<li><a class="noeffect"><span class="name">Shrimp</span></a></li>
		<li><a class="noeffect"><span class="name">Squash, summer</span></a></li>
		<li><a class="noeffect"><span class="name">teapot</span></a></li>
		<li><a class="noeffect"><span class="name">toothbrush</span></a></li>
		<li><a class="noeffect"><span class="name">ham</span></a></li>
		<li class="hidden autolisttext"><a class="noeffect" href="#">Load 10 
		more items...</a></li>
	</ul>
</div>
<div id="footer">
	<a href="http://iwebkit.net">Powered by iWebKit</a></div>
</body>
</html>
