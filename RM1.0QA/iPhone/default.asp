<%@ Language=VBScript %>
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   strUser     = Request.Cookies("iPhone-User")
   strPassword = Request.Cookies("iPhone-Password")
   
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta content="yes" name="apple-mobile-web-app-capable" />
<meta content="text/html; charset=iso-8859-1" http-equiv="Content-Type" />
<meta content="minimum-scale=1.0, width=device-width, maximum-scale=0.6667, user-scalable=no" name="viewport" />
<link href="css/style.css" rel="stylesheet" type="text/css" />
<script src="javascript/functions.js" type="text/javascript"></script>
<title>Rate-Monitor for iPhone</title>
<meta content="keyword1,keyword2,keyword3" name="keywords" />
<meta content="Description of your site" name="description" />
<script type="text/javascript" >
var BrowserDetect = {
	init: function () {
		this.browser = this.searchString(this.dataBrowser) || "An unknown browser";
		this.version = this.searchVersion(navigator.userAgent)
			|| this.searchVersion(navigator.appVersion)
			|| "an unknown version";
		this.OS = this.searchString(this.dataOS) || "an unknown OS";
	},
	searchString: function (data) {
		for (var i=0;i<data.length;i++)	{
			var dataString = data[i].string;
			var dataProp = data[i].prop;
			this.versionSearchString = data[i].versionSearch || data[i].identity;
			if (dataString) {
				if (dataString.indexOf(data[i].subString) != -1)
					return data[i].identity;
			}
			else if (dataProp)
				return data[i].identity;
		}
	},
	searchVersion: function (dataString) {
		var index = dataString.indexOf(this.versionSearchString);
		if (index == -1) return;
		return parseFloat(dataString.substring(index+this.versionSearchString.length+1));
	},
	dataBrowser: [
		{
			string: navigator.userAgent,
			subString: "Chrome",
			identity: "Chrome"
		},
		{ 	string: navigator.userAgent,
			subString: "OmniWeb",
			versionSearch: "OmniWeb/",
			identity: "OmniWeb"
		},
		{
			string: navigator.vendor,
			subString: "Apple",
			identity: "Safari",
			versionSearch: "Version"
		},
		{
			prop: window.opera,
			identity: "Opera"
		},
		{
			string: navigator.vendor,
			subString: "iCab",
			identity: "iCab"
		},
		{
			string: navigator.vendor,
			subString: "KDE",
			identity: "Konqueror"
		},
		{
			string: navigator.userAgent,
			subString: "Firefox",
			identity: "Firefox"
		},
		{
			string: navigator.vendor,
			subString: "Camino",
			identity: "Camino"
		},
		{		// for newer Netscapes (6+)
			string: navigator.userAgent,
			subString: "Netscape",
			identity: "Netscape"
		},
		{
			string: navigator.userAgent,
			subString: "MSIE",
			identity: "Explorer",
			versionSearch: "MSIE"
		},
		{
			string: navigator.userAgent,
			subString: "Gecko",
			identity: "Mozilla",
			versionSearch: "rv"
		},
		{ 		// for older Netscapes (4-)
			string: navigator.userAgent,
			subString: "Mozilla",
			identity: "Netscape",
			versionSearch: "Mozilla"
		}
	],
	dataOS : [
		{
			string: navigator.platform,
			subString: "Win",
			identity: "Windows"
		},
		{
			string: navigator.platform,
			subString: "Mac",
			identity: "Mac"
		},
		{
			   string: navigator.userAgent,
			   subString: "iPhone",
			   identity: "an iPhone or iPod"
	    },
		{
			string: navigator.platform,
			subString: "Linux",
			identity: "Linux"
		}
	]

};
BrowserDetect.init();
</script>
</head>
<body>


<div id="topbar">
	<div id="title">
		Rate-Monitor</div>
	<div id="leftbutton">
		<a href="http://www.rate-monitor.com">PC website</a></div>
</div>
<div id="content">
	<ul class="pageitem">
		<li class="textbox"><span class="header">Welcome</span><p>Welcome to the 
		Rate-Monitor site! You're using
<script type="text/javascript">
  document.write(BrowserDetect.browser);
  document.write(" on " + BrowserDetect.OS);
</script>
 
  </p>
		</li>
		<li class="menu"><a href="changelog.html">
		<img alt="changelog" src="thumbs/start.png" /><span class="name">What&#39;s 
		New?</span><span class="arrow"></span></a></li>
	</ul>
	
	<span class="graytitle">Please Login</span>
<form name="login" action="login.asp" method="post">
	<ul class="pageitem">
		<li class="form"><input name="user"     placeholder="Username" type="text" value="<%=strUser %>"/></li>
		<li class="form"><input name="password" placeholder="Password" type="password" value="<%=strPassword %>" /></li>
		<li class="form"><input name="input Button" type="submit" value="Login" /></li>
	</ul>
</form>	
</div>

<div id="footer">
	<a class="noeffect" href="http://www.rate-highway.com">Powered by Rate-Highway, Inc.</a></div>

</body>

</html>
