<%@ Language=VBScript %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Highway, Inc.| Rate-Monitor Upload Monitor</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}


function CheckLoginAlerts()
{
	if("ANY"=="ANY") 
			{
			alert("No alerts at this time.");  
			return true ;
			}

}

//-->
</script>
<link rel="stylesheet" type="text/css" href="../inc/rh_standard.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="" style="">
<!--
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="" style="filter:progid:DXImageTransform.Microsoft.Gradient(endColorstr='#33444D', startColorstr='#FFFFFF', gradientType='0');">
-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="../images/top_tile.gif"><img src="../images/top.jpg"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="../images/b_tile.gif">
<table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr>
         <td><img src="../images/b_left.jpg"></td>
          <td>
          <img src="../images/blanks/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></td>
          <td>
          <img src="../images/blanks/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></td>
          <td>
          <img src="../images/blanks/b_search_cri_of.gif" name="s3" border="0" id="s3"></td>
          <td>
          <img src="../images/blanks/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></td>
          <td>
          <img src="../images/blanks/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></td>
          <td>
          <img src="../images/blanks/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></td>
          <td>
          <img src="../images/blanks/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="../images/med_bar_tile.gif"><img src="../images/med_bar.gif"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="../images/user_tile.gif">
<table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../images/user_left.gif" width="580" height="31"></td>
          <td background="../images/user_tile.gif">
<table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td valign="bottom">
<table width="100" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td><div align="right"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"></div></td>
                    </tr>
                    <tr>
                      <td><img src="../images/separator.gif" width="183" height="6"></td>
                    </tr>
                  </table>
                </td>
                <td><img src="../images/user_tile.gif"></td>
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
    <td background="../images/h_tile.gif"><table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="../images/nothing_top.gif"></td>
          <td><img src="../images/h_right.gif"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<form method="POST" action="../system_login.asp" name="login" OnSubmit="CheckLoginAlerts();return true" class="login">
<table width="770" border="0" cellspacing="0" cellpadding="2">
  <tr> 
    <td width="20"><img src="../images/separator.gif"></td>
    <td width="858" valign="top"> <p><font size="4">Rate-Highway Data Upload 
    Monitor</font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
      <p><font color="#800000"><b><a name="Top">Data Upload Help</a></b></font><br>
        <font face="Arial, Helvetica, sans-serif" size="2">The Rate-Highway data 
      upload program is used to send utilization information from your counter 
      system to this Rate-Monitor system.</font></p>
    <p><font face="Arial, Helvetica, sans-serif" size="2">This page discusses:<br>
    <a href="#Modify_Settings">Modifying settings</a><br>
    <a href="#Fields">The Upload fields</a><br>
    <a href="#Installing">Installing</a> (including links to download the 
	software)</font></p>
    <p><font face="Arial, Helvetica, sans-serif" size="2">(Download help located
    <a href="../download/default.asp">here</a>)</font></p>
      <p><font size="2">If you have any questions about this program or the 
      Rate-Monitor system, please email
      <a href="mailto:support@rate-highway.com">support@rate-highway.com</a> for 
      answers to all your questions.</font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0" id="table2">
        <tr> 
          <td><img src="../images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
    <p><b><font face="Arial, Helvetica, sans-serif"><a name="Modify_Settings">Modify Settings</a></font></b></p>
    <p><font face="Arial, Helvetica, sans-serif" size="2">To modify the settings 
    on the Rate-Highway data upload program, once the program is installed, it 
    will reside in your system tray (the section of your computer - usually in 
    the bottom right hand side of your screen that has the clock and other small 
    icons like Outlook and anti-virus, etc.) on the counter system server. Look for an 
    icon that looks like this:</font></p>
    <table border="0" style="border-collapse: collapse" width="100%" id="table3">
      <tr>
        <td>&nbsp;</td>
        <td width="347">
        <p align="center">
	<img border="0" src="../images/upload_icon_large.gif" width="102" height="103" alt="Rate-Highway ftp upload tray icon"></td>
        <td>&nbsp;</td>
        <td>
        <p align="center">
	<img border="0" src="../images/upload%20context%20menu.JPG" width="192" height="144" alt="Select the Options menu item to show the options window"></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td width="347">
        <p align="center"><font size="2">1. Right-click the Rate-Highway Upload 
        Icon</font></td>
        <td>&nbsp;</td>
        <td>
        <p align="center"><font size="2">2. Select the &quot;Options&quot; menu item</font></td>
        <td>&nbsp;</td>
      </tr>
    </table>
    <p><font face="Arial, Helvetica, sans-serif" size="2">Right-click the icon 
    in your system tray 
    and select the &quot;options&quot; menu item. This will invoke the window discussed 
    below.</font></p>
    <p><font face="Arial, Helvetica, sans-serif" size="2"><a href="#Top">Return 
    to top</a></font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0" id="table1">
        <tr> 
          <td><img src="../images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
	<p><strong><font face="Arial, Helvetica, sans-serif"><a name="Fields">Fields</a></font></strong></p>
	<ol>
		<li><font face="Arial, Helvetica, sans-serif" size="2"><b>Directory to 
		monitor</b> - This field holds the name of the directory on your counter 
        system server that contains the utilization files that will be uploaded 
        to Rate-Monitor. Your counter system vendor will tell you which 
        directory to check as it may be custom on your machine.<br>
&nbsp;</font></li>
        <li><font face="Arial, Helvetica, sans-serif" size="2"><b>Destination URL</b> - 
        This is the name of the server in which to send the information. Please 
        make sure it is set to
        </font><font face="Courier New" size="2">ftp://zeus.rate-monitor.com<br>
&nbsp;</font></li>
        <li><font face="Arial, Helvetica, sans-serif" size="2"><b>Customer number</b> - 
        This is your unique customer number assigned to you by our support 
        staff. If you have misplaced your customer number please click here to
        <a href="mailto:support@rate-highway.com?subject=Please send me my customer number">
        email support</a> <br>
&nbsp;</font></li>
        <li><font face="Arial, Helvetica, sans-serif" size="2"><b>Upload interval</b> - 
        The amount of time that should pass before the system checks for new 
        files to upload. This should be set to &quot;1 Hour&quot; unless you are told to 
        do otherwise.</font></li>
	</ol>
	<p>&nbsp;</p>
	<p align="center">
	<img border="0" src="../images/upload.jpg" width="440" height="520" alt="Rate-Highway ftp upload configuration"></p>
    <p>&nbsp;</p>
      <p><font size="2" face="Arial, Helvetica, sans-serif"> Clicking on the OK 
      button will save any changes you have made and close the window, clicking 
      Cancel will not save any changes you made and also close the window. </font></p>
    <p><font face="Arial, Helvetica, sans-serif" size="2"><a href="#Top">Return 
    to top</a></font></p>
      <table width="745" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="../images/ruler.gif" width="745" height="2"></td>
        </tr>
      </table>
	<p><strong><font face="Arial, Helvetica, sans-serif"><a name="Installing">Installing</a></font></strong></p>
	<p><strong style="font-weight: 400; font-style: italic">
	<font face="Arial, Helvetica, sans-serif" size="2">NOTE: This program 
	requires Microsoft .Net 2.0 to be installed. If you do not have it currently 
	installed, you can easily do so by visiting this link to the
	<a target="_blank" href="http://www.microsoft.com/downloads/details.aspx?FamilyID=0856eacb-4362-4b0d-8edd-aab15c5e04f5&DisplayLang=en">
	Microsoft Download</a> service.</font></strong></p>
	<p>1<font size="2">. To install this program click
      <a href="Rate-Highway%20File%20Transfer.msi">here</a>. You will then be 
    presented with the following pop-up window:</font></p>
    <p align="center">
	<img border="0" src="../images/security_warning.jpg" alt="Rate-Highway ftp upload configuration"></p>
    <p>2<font size="2">. Click run to install, and select all the default options (make 
      sure you are doing this on only your counter system server). You do not 
    need to install this on any of your workstations or regular computers.</font></p>
    <p align="center">
	<img border="0" src="../images/security_warning2.jpg" alt="Rate-Highway ftp upload configuration" width="465" height="231"></p>
    <p>3<font size="2">. Again click run to install, and next the regular 
    install program will begin. You may use all the default settings. Once 
    installed you should reboot your machine and from this point forward, the 
    Upload program icon should appear in your system tray.</font></p>
    <p align="center">&nbsp;</p>
    <p><font face="Arial, Helvetica, sans-serif" size="2"><a href="#Top">Return 
    to top</a></font></p></td>
  </tr>
</table>
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>