<%@ Language=VBScript %>


<% 	Response.Expires = -1
   	Response.cachecontrol="private" 
   	Response.AddHeader "pragma", "no-cache"

	Rem Clear the bastard out...  
	Response.Cookies("rate-monitor.com") = ""
	
	Dim strLogin 

	Rem Retrieve the login id but clerar the rest out, in case
	strLogin = Request.Cookies("rate-monitor.com")("LoginCode")


	Response.Cookies("rate-monitor.com")("password") = ""
	Response.Cookies("rate-monitor.com")("remember Me") = ""
	Response.Cookies("rate-monitor.com")("testing") = ""
	Response.Cookies("rate-monitor.com")("user_id") = ""
	Response.Cookies("rate-monitor.com")("user_name") = ""
	Response.Cookies("rate-monitor.com")("client_userid") = ""
	Response.Cookies("rate-monitor.com")("vend_cd") = ""		


	Dim strDestURL
	
	

	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Highway, Inc. | Welcome and please login</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">

<script language="JavaScript" type="text/JavaScript">
function jumpTo()
{

	var NewAction
	var SelIndex
	
	SelIndex = document.login.system.selectedIndex
	NewAction = document.login.system.options[selectedIndex].text

	document.login.action = NewAction

}

function CheckLoginAlerts()
{
	if("ANY"!="ANY") 
			{
			//alert("No alerts at this time.");  
			return true ;
			}
	return true;
}

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
//-->
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="document.login.email_address.focus()"
style="filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#33444D', startColorstr='#FFFFFF', gradientType='0');">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif"><img src="images/top.jpg"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/b_tile.gif">
<table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/b_left.jpg"></td>
          <td>
          <img src="images/blanks/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></td>
          <td>
          <img src="images/blanks/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></td>
          <td>
          <img src="images/blanks/b_search_cri_of.gif" name="s3" border="0" id="s3"></td>
          <td>
          <img src="images/blanks/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></td>
          <td>
          <img src="images/blanks/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></td>
          <td>
          <img src="images/blanks/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></td>
          <td>
          <img src="images/blanks/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/user_tile.gif">
<table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/user_left.gif" width="580" height="31"></td>
          <td background="images/user_tile.gif">
<table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="bottom">
<table width="100" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><div align="right"><a href="report_login.asp"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Not Logged In</a></div></td>
                    </tr>
                    <tr>
                      <td><img src="images/separator.gif" width="183" height="6"></td>
                    </tr>
                  </table>
                </td>
                <td><img src="images/user_tile.gif"></td>
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
          <td><img src="images/nothing_top.gif"></td>
          <td><img src="images/h_right.gif"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<form method="POST" action="system_login.asp" name="login" OnSubmit="CheckLoginAlerts();return true" class="login">
<table width="770" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="858" valign="top"> <p>&nbsp;</p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p><b><font color="#800000">Welcome to Rate-Monitor -- Please login to 
      view your report</font></b><br>
        <font size="2" face="Arial, Helvetica, sans-serif">To login, please enter 
      your username (or email address) and password, then click the Log In button. 
        Please remember that your password is case sensitive. </font> </p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/alt_color.gif">
            <img src="images/separator.gif" width="745" height="15">
          </td>
        </tr>
      </table>
      <table width="745" border="0" cellpadding="0" cellspacing="0" background="images/alt_color.gif">
        <tr valign="bottom"> 
          <td width="163"><font size="2" face="Arial, Helvetica, sans-serif">
          <img src="images/ti_log_in.gif"></font></td>
          <td width="104">
          <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;User 
            Name:</font> 
          <td width="182">
          <p align="left"><font size="2" face="Arial, Helvetica, sans-serif"> 
            <input type="text" name="email_address" size="20" tabindex="1" onfocus="this.className='focus';cl(this,'email');" onblur="this.className='';fl(this,'email');" style="width: 150" value="<%=strLogin %>"></font><td width="296"> <font size="2" face="Arial, Helvetica, sans-serif" > 
            <input name="login" type="submit" id="Open2224" value="      Log In      " tabindex="3" class="rh_button">&nbsp; </font></td>
        </tr>
      </table>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td background="images/alt_color.gif"><table width="745" border="0" cellpadding="0" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom"> 
                <td width="163"><font size="2" face="Arial, Helvetica, sans-serif"><img src="images/separator.gif" width="162" height="20"></font></td>
                <td width="104">
                <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Password:</font> 
                <td width="182"><font size="2" face="Arial, Helvetica, sans-serif"> 
                  <input type="password" name="password" size="20" tabindex="2" onfocus="this.className='focus';cl(this,'');" onblur="this.className='';fl(this,'');" style="width: 150"></font><td width="296"> <font size="2" face="Arial, Helvetica, sans-serif"> 
                  &nbsp;<a href="mailto:support@rate-highway.com?subject=Password email request">email password</a></font></td>
              </tr>
              <tr valign="bottom"> 
                <td width="163">&nbsp;</td>
                <td width="104">
                <!--
                <font size="2" face="Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;System:</font> 
                -->
                <td width="182">
                <!--
                <select size="1" name="system" style="width: 150" >
                <option value="hotel">Hotel - Offline</option>
	            <option selected value="car">Car - Online</option>
                <option value="air">Air - Online</option>
                <option value="cruise">Cruise - Offline</option>
                </select>
                -->
                <td width="296"> <font size="2" face="Arial, Helvetica, sans-serif"> 
                  &nbsp;<a href="javascript:centerPopUp( 'change_password.asp', 'change', 620, 300 )">change password</a></font></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
        <tr valign="bottom"> 
          <td width="161">
          <p align="center"><img src="images/separator.gif" width="30" height="15"></td>
        </tr>
      </table>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
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
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
  </tr>
</table>
<input type="hidden" name="request_url" value="<%=Session("request_url") %>">
<input type="hidden" name="fromform" value="report_login">
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<script language="javascript">
	document.login.email_address.focus();
</script>