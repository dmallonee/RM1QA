<%@ Language=VBScript %>
<% 	Response.Expires = -1
   	Response.cachecontrol="private" 
   	Response.AddHeader "pragma", "no-cache"

	Dim strLogin 

	Rem Retrieve the login id but clerar the rest out, in case
	strLogin = Request.Cookies("rate-monitor.com")("LoginCode")

	Rem Clear the bastard out...  
	Response.Cookies("rate-monitor.com") = ""

	Response.Cookies("rate-monitor.com")("password") = ""
	Response.Cookies("rate-monitor.com")("remember Me") = ""
	Response.Cookies("rate-monitor.com")("testing") = ""
	Response.Cookies("rate-monitor.com")("user_id") = ""
	Response.Cookies("rate-monitor.com")("user_name") = ""
	Response.Cookies("rate-monitor.com")("client_userid") = ""
	Response.Cookies("rate-monitor.com")("vend_cd") = ""	
	

   If Len(strLogin) = 0 Then 
      strControlToSelect = "document.login.email_address.focus()"
   Else 
      strControlToSelect = "document.login.password.focus()"
   End If 
	
%>
<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN'>
<html>
<head>
<meta http-equiv='Content-Language' content='en-us'>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="SHORTCUT ICON" href="favicon.ico?v=2">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<title>Rate Monitor | Rate Automation made easy</title>
<script language="javascript" type="text/javascript" src="inc/sitewide.js" ></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function centerPopUp( url, name, width, height, scrollbars ) { 
 
	if( scrollbars == null ) scrollbars = "0" 
 
	str  = ""; 
	str += "resizable=1,"; 
	str += "scrollbars=" + scrollbars + ","; 
	str += "width=" + width + ","; 
	str += "height=" + height + ","; 
 
	if ( window.screen ) { 
		var ah = screen.availHeight - 30; 
		var aw = screen.availWidth - 10; 
 
		var xc = ( aw - width ) / 2; 
		var yc = ( ah - height ) / 2; 
 
		str += ",left=" + xc + ",screenX=" + xc; 
		str += ",top=" + yc + ",screenY=" + yc; 
	} 
	window.open( url, name, str ); 
} 

function checkform() {
	var why = "";
	why += checkUsername(document.login_form.email_address.value);
	why += checkPassword(document.login_form.password.value);

	if (why != "") {
		alert(why);
		return false;
	}
	return true;
}

function checkPassword(strng) {
	var error = "";
	if(strng == "") {
		error = "Please enter your password.\n";
		login_form.password.style.background = 'rgba(162,0,0,0.5)'; 
	}
	else {
		login_form.password.style.background = 'White';
	}
	return error;
}
function checkUsername(strng) {
	var error = "";
	if(strng == "") {
		error = "Please enter your user name.\n";
		login_form.email_address.style.background = 'rgba(162,0,0,0.5)'; 
	}
	else {
		login_form.email_address.style.background = 'White';
	}
	return error;
}
-->
//-->
</script>
</head>
   
<body style="background-color:#FFFFFF;margin-left: 0px;margin-top: 0px;filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#33444D', startColorstr='#FFFFFF', gradientType='0');" onload="<%=strControlToSelect %>" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td style="background-image:url('images/top_tile.gif')">
    <a target="_blank" href="http://www.rate-highway.com">
    <img src="images/top.jpg" width="770" height="91" border="0" alt="Visit Rate-Highway" ></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td style="background-image:url('images/b_tile.gif')">
<table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img alt='' src="images/b_left.jpg"></td>
          <td>
          <img alt='' src="images/blanks/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_search_cri_of.gif" name="s3" border="0" id="s3"></td>
          <td>
          <img alt='' src="images/blanks/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></td>
          <td>
          <img alt='' src="images/blanks/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img alt='' src="images/med_bar.gif"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/user_tile.gif">
<table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img alt='' src="images/user_left.gif" width="580" height="31"></td>
          <td background="images/user_tile.gif">
<table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="bottom">
<table width="100" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><div align="right"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Not Logged In</font></div></td>
                    </tr>
                    <tr>
                      <td><img alt='' src="images/separator.gif" width="183" height="6"></td>
                    </tr>
                  </table>
                </td>
                <td><img alt='' src="images/user_tile.gif"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/h_tile.gif"><table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img alt='' src="images/nothing_top.gif"></td>
          <td><img alt='' src="images/h_right.gif"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<form method="POST" action="system_login.asp" name="login_form" OnSubmit="return checkform();" class="login">
<table width="770" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="858" valign="top"> 
    <p align="center"><span style="font-family: Arial; font-weight: 700">
    <font size="4"><br>
    </font></span><span style="font-family: Arial; font-weight: 700"><b>Rate Monitor</b><font color="#800000"> 
    - rate automation based on intelligent human input.</font></span></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img alt='' src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p><b><font color="#800000">Welcome, please login</font></b><br>
        <font size="2" face="Arial, Helvetica, sans-serif">To login, please enter 
      your username (or email address) and password, then click the Log In button. 
        Please remember that your password is case sensitive. </font> </p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img alt='' src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/alt_color.gif">
            <img alt='' src="images/separator.gif" width="745" height="15">
          </td>
        </tr>
      </table>
      <table width="745" border="0" cellpadding="0" cellspacing="0" background="images/alt_color.gif">
        <tr valign="bottom"> 
          <td width="163"><font size="2" face="Arial, Helvetica, sans-serif">
          <img alt='' src="images/ti_log_in.gif"></font></td>
          <td width="104" align="left" >
          <font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;User 
            Name:</font> 
          <td width="182" align="left" >
          <font size="2" face="Arial, Helvetica, sans-serif"> 
            <input type="text" name="email_address" size="20" tabindex="1" onfocus="this.className='focus';cl(this,'email');" onblur="this.className='';fl(this,'email');" style="width: 150" value="<%=strLogin %>"></font><td width="296"> <font size="2" face="Arial, Helvetica, sans-serif" > 
            <input name="login" type="submit" id="Open2224" value="      Log In      " tabindex="3" class="rh_button">&nbsp; </font></td>
        </tr>
      </table>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/alt_color.gif"><table width="745" border="0" cellpadding="0" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom"> 
                <td width="163"><font size="2" face="Arial, Helvetica, sans-serif"><img alt='' src="images/separator.gif" width="162" height="20"></font></td>
                <td width="104" align="left">
                <font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Password:</font> 
                <td width="182"><font size="2" face="Arial, Helvetica, sans-serif"> 
                  <input type="password" name="password" size="20" tabindex="2" onfocus="this.className='focus';cl(this,'');" onblur="this.className='';fl(this,'');" style="width: 150"></font><td width="296"> <font size="2" face="Arial, Helvetica, sans-serif"> 
                  &nbsp;<a href="mailto:support@rate-highway.com?subject=Password email request">email password</a></font></td>
              </tr>
              <tr valign="bottom"> 
                <td width="163">&nbsp;</td>
                <td width="104">&nbsp;</td>
                <td width="182">&nbsp;</td>
                <td width="296"> <font size="2" face="Arial, Helvetica, sans-serif"> 
                  &nbsp;<a href="javascript:centerPopUp( 'change_password.asp', 'change', 620, 300 )">change password</a></font></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
        <tr valign="bottom"> 
          <td width="161">
          <p align="center"><img alt='' src="images/separator.gif" width="30" height="15"></td>
        </tr>
      </table>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img alt='' src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p><font size="3" face="Arial, Helvetica, sans-serif"><br>
        <strong>Forgot your password? </strong> </font></p>
      <p><font size="2" face="Arial, Helvetica, sans-serif"> If you have forgotten 
        your password, enter your username and click the &quot;email password&quot; 
      link. 
        Your password will be emailed to the address stored with your account. 
        So, make sure you keep your email address up to date in our records. </font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img alt='' src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p><font size="3" face="Arial, Helvetica, sans-serif"><br>
           <a href="https://www.mozilla.org/en-US/firefox/new/" target="_blank"><img src="images/firefox.png" title="Firefox" align="right" height="64" style="padding-right:15px;border:0;"/></a>
           <a href="http://windows.microsoft.com/en-us/internet-explorer/download-ie" target="_blank"><img src="images/ie.png" title="Internet Explorer" align="right" height="64" style="border:0;"/></a>
           <a href="http://support.apple.com/downloads/#safari" target="_blank"><img src="images/safari.png" title="Safari" align="right" height="64" style="padding-left:15px;border:0;"/></a>
        <strong>Recommended Browsers </strong> </font></p>
      <p><font size="2" face="Arial, Helvetica, sans-serif"> 
        Rate-Monitor is designed for best use with recent versions of Safari (5.0+), Internet Explorer (8.0+), and Firefox (12.0+).
            Please be sure to use one of these browsers when visiting the Rate-Monitor website.
		            You can upgrade your browser by clicking on one of the links to the right.<br />
         </font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img alt='' src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
  </tr>
</table>
</form>
<!--#INCLUDE FILE="footer.asp"-->
<script language="javascript" type="text/javascript" src="https://seal.godaddy.com/getSeal?sealID=14250815080196bac4127102cc7738dc17cf907881377017190090109"></script>
</p>
<% If strLogin = "" Then %> 
<script language="javascript" type="text/javascript" >
	document.login.email_address.focus();
</script>
<% Else %>
<script language="javascript" type="text/javascript" >
	document.login.password.focus();
</script>
<% End If %>
</body>
</html>
