<%@ Language=VBScript %>


<% 	Response.Expires = -1
   	Response.cachecontrol="private" 
   	Response.AddHeader "pragma", "no-cache"

	Rem Clear the bastard out...  
	Response.Cookies("rate-monitor.com") = ""
	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Account Suspended</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif');document.login.email_address.focus()"
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
                      <td><div align="right"><a href="suspended.asp"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Not Logged In</a></div></td>
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
<form method="POST" action="system_login.asp" name="login" class="login">
<table border="0" cellspacing="0" cellpadding="0" align="center" style="width: 748px">
  <tr> 
    <td width="858" valign="top"> <p>&nbsp;</p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p><font color="#800000"><b>Account Suspended</b></font><br>
        </p>
    <p><font size="2" face="Arial, Helvetica, sans-serif">Your account is not 
    currently enabled. Please contact your Rate-Highway account representative 
    for more information. </font> </p>
    <p>&nbsp;</p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
  </tr>
</table>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>

