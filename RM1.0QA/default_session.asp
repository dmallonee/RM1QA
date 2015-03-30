<%@ Language=VBScript %>


<% 	Response.Expires = -1
   	Response.cachecontrol="private" 
   	Response.AddHeader "pragma", "no-cache"


	Dim strLogin 

	Rem Retrieve the login id but clerar the rest out, in case
	strLogin = Request.Cookies("rate-monitor.com")("LoginCode")


	Rem Clear the bastard out...  
	Response.Cookies("rate-monitor.com") = ""
	
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor.com | Welcome and please login</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
</head>
<% If Len(strLogin) = 0 Then %>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="document.login_form.email_address.focus()"
<% Else %>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="document.login_form.password.focus()"
<% End If %>
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
          <a href="search_profiles_car.asp" onMouseOver="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/blanks/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td>
          <td><a href="search_queue_car.asp" onMouseOver="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/blanks/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
          <td><a href="search_criteria_car.asp" onMouseOver="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/blanks/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('ra','','images/b_rate_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/blanks/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('al','','images/b_alert_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/blanks/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('us','','images/b_user_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/blanks/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
          <td>
          <a onMouseOver="MM_swapImage('sy','','images/b_system_on.gif',1)" onMouseOut="MM_swapImgRestore()" href="javascript:not_enabled()">
          <img src="images/blanks/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td>
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
                      <td><div align="right"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Not Logged In</div></td>
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
<form method="POST" action="system_login.asp" name="login_form" OnSubmit="CheckLoginAlerts();return true" class="login">
<div align="center">
<table width="770" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td width="20"><img src="images/separator.gif"></td>
    <td width="858" valign="top"> <p>&nbsp;</p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p><b><font color="#800000">Your session has expired. Please login again</font></b><br>
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
            <input type="text" name="email_address" size="20" tabindex="1" value="<%=strLogin %>" onfocus="this.className='focus';cl(this,'email');" onblur="this.className='';fl(this,'email');" style="width: 150" ></font><td width="296"> <font size="2" face="Arial, Helvetica, sans-serif"> 
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
                  <input type="password" name="password" size="20" tabindex="2" onfocus="this.className='focus';cl(this,'');" onblur="this.className='';fl(this,'');" style="width: 150" ></font><td width="296"> <font size="2" face="Arial, Helvetica, sans-serif"> 
                  &nbsp;
          </font></td>
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
        &nbsp;</font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p>&nbsp;</p></td>
  </tr>
</table>
</div>
</form>
<!--#INCLUDE FILE="footer.asp"-->
<p align="center">
<script language="Javascript" src="https://seal.godaddy.com/getSeal?sealID=14250815080196bac4127102cc7738dc17cf907881377017190090109"></script>
</p>

<p>&nbsp;</p>
</body>
</html>
<% If strLogin = "" Then %> 
<script language="javascript">
    document.login_form.email_address.focus();
</script>
<% Else %>
<script language="javascript">
    document.login_form.password.focus();
</script>
<% End If %>