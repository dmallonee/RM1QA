<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" --> 
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 

<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 0
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">

<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; System</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
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


function openWindow(theURL,winName,features) 
{ //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>

<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style>
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
-->
</style>
<base target="_self">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif"><img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/b_tile.gif">
<table width="400" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/b_left.jpg" width="62" height="32"></td>
          <td><a href="search_profiles_car.asp" onMouseOver="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td>
          <td><a href="search_queue_car.asp" onMouseOver="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
          <td><a href="search_criteria_car.asp" onMouseOver="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('ra','','images/b_rate_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('al','','images/b_alert_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
          <td><a href="javascript:not_enabled()" onMouseOver="MM_swapImage('us','','images/b_user_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
          <td><a href="system_locations_car.asp" onMouseOver="MM_swapImage('sy','','images/b_system_on.gif',1)" onMouseOut="MM_swapImgRestore()">
          <img src="images/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif" width="12" height="8"></td>
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
                      <td><div align="right">
                  <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div></td>
                    </tr>
                    <tr>
                      <td><img src="images/separator.gif" width="183" height="6"></td>
                    </tr>
                  </table>
                </td>
                <td><img src="images/user_tile.gif" width="7" height="31"></td>
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
          <td><img src="images/h_system.gif" width="368" height="31"></td>
          <td><img src="images/h_right.gif" width="402" height="31"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;&nbsp;&nbsp; <br>
&nbsp;<font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;<a href="javascript:not_enabled()">[custom 
city codes]</a>&nbsp; <b>[system status]</b> </font><br>
&nbsp;</p>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" cellspacing="0" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
  <table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" id="profiles">
    <tr>
      <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="30">&nbsp;</td>
      <td class="profile_header" width="66"  height="45">Selected</td>
      <td class="profile_header" width="109" height="45">Machine</td>
      <td class="profile_header" width="83" height="45">Active Searches</td>
      <td class="profile_header" width="98" height="45">Proxy<br>
      Group</td>
      <td class="profile_header" width="76" height="45">Searches Today</td>
      <td class="profile_header" width="105" height="45">Machine Rating (RPH)</td>
      <td class="profile_header" width="127" height="45">Rating Index</td>
      <td class="profile_header" width="94" height="45">Comparative Index</td>
      <td class="profile_header" width="97" height="45">Change<br>
      + / -</td>
      <td class="profile_header" width="126" height="45">Status</td>
      <td class="profile_header" width="635" height="45">Notes</td>
    </tr>
    <tr>
      <td width="26" class="profile_light" height="20">1</td>
      <td width="70" bgcolor="#FDC677" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="109" class="profile_light" height="20">
      <a href="javascript:void(0)" onClick="openWindow('http://<%=Request.Servervariables("REMOTE_ADDR")%>/tsweb?Server=64.147.13.100','SearchMachine','scrollbars=no,status=no,resize=no,width=820,height=620,top=100,left=100,toolbar=no,menubar=no,location=no')">Search 1</a></td>
      <td width="83" class="profile_light" height="20">0</td>
      <td width="98" class="profile_light" height="20">1</td>
      <td width="76" class="profile_light" height="20">4,673</td>
      <td width="105" class="profile_light" height="20">1000</td>
      <td width="127" class="profile_light" height="20">131%</td>
      <td width="94" class="profile_light" height="20">100%</td>
      <td width="97" class="profile_light" height="20">0</td>
      <td width="126" class="profile_light" height="20">idle - C</td>
      <td width="635" class="profile_light" height="20">Primary demo</td>
    </tr>
    <tr>
      <td width="26" class="profile_dark" height="20">2</td>
      <td width="70" bgcolor="#E07D1A" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="109" class="profile_dark" height="20">
      <a target="_blank" href="http://64.147.13.101/tsweb">Search 2</a></td>
      <td width="83" class="profile_dark" height="20">0</td>
      <td width="98" class="profile_dark" height="20">1</td>
      <td width="76" class="profile_dark" height="20">0</td>
      <td width="105" class="profile_dark" height="20">1000</td>
      <td width="127" class="profile_dark" height="20">N/A</td>
      <td width="94" class="profile_dark" height="20">N/A</td>
      <td width="97" class="profile_dark" height="20">&nbsp;</td>
      <td width="126" class="profile_dark" height="20">off-line</td>
      <td width="635" class="profile_dark" height="20">&nbsp;</td>
    </tr>
    <tr>
      <td width="26" class="profile_light" height="20">3</td>
      <td width="70" bgcolor="#FDC677" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="109" class="profile_light" height="20">
      <a href="http://64.147.13.101/tsweb">Search 3</a></td>
      <td width="83" class="profile_light" height="20">0</td>
      <td width="98" class="profile_light" height="20">1</td>
      <td width="76" class="profile_light" height="20">0</td>
      <td width="105" class="profile_light" height="20">1000</td>
      <td width="127" class="profile_light" height="20">N/A</td>
      <td width="94" class="profile_light" height="20">N/A</td>
      <td width="97" class="profile_light" height="20">&nbsp;</td>
      <td width="126" class="profile_light" height="20">off-line</td>
      <td width="635" class="profile_light" height="20">&nbsp;</td>
    </tr>
    <tr>
      <td width="26" class="profile_dark" height="20">4</td>
      <td width="70" bgcolor="#E07D1A" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="109" class="profile_dark" height="20">
      <a href="http://64.147.13.101/tsweb">Search 4</a></td>
      <td width="83" class="profile_dark" height="20">0</td>
      <td width="98" class="profile_dark" height="20">2</td>
      <td width="76" class="profile_dark" height="20">0</td>
      <td width="105" class="profile_dark" height="20">1000</td>
      <td width="127" class="profile_dark" height="20">N/A</td>
      <td width="94" class="profile_dark" height="20">N/A</td>
      <td width="97" class="profile_dark" height="20">&nbsp;</td>
      <td width="126" class="profile_dark" height="20">idle</td>
      <td width="635" class="profile_dark" height="20">Overlandwest dedicated</td>
    </tr>
    <tr>
      <td width="26" class="profile_light" height="20">5</td>
      <td width="70" bgcolor="#FDC677" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="109" class="profile_light" height="20">
      <a href="http://64.147.13.101/tsweb">Search 5</a></td>
      <td width="83" class="profile_light" height="20">0</td>
      <td width="98" class="profile_light" height="20">2</td>
      <td width="76" class="profile_light" height="20">0</td>
      <td width="105" class="profile_light" height="20">1000</td>
      <td width="127" class="profile_light" height="20">N/A</td>
      <td width="94" class="profile_light" height="20">N/A</td>
      <td width="97" class="profile_light" height="20">&nbsp;</td>
      <td width="126" class="profile_light" height="20">purge</td>
      <td width="635" class="profile_light" height="20">In maintenance</td>
    </tr>
    <tr>
      <td width="26" class="profile_dark" height="20">6</td>
      <td width="70" bgcolor="#E07D1A" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="109" class="profile_dark" height="20">
      <a href="http://64.147.13.101/tsweb">Search 6</a></td>
      <td width="83" class="profile_dark" height="20">0</td>
      <td width="98" class="profile_dark" height="20">2</td>
      <td width="76" class="profile_dark" height="20">1,452</td>
      <td width="105" class="profile_dark" height="20">1000</td>
      <td width="127" class="profile_dark" height="20">129%</td>
      <td width="94" class="profile_dark" height="20">98%</td>
      <td width="97" class="profile_dark" height="20">(2%)</td>
      <td width="126" class="profile_dark" height="20">idle</td>
      <td width="635" class="profile_dark" height="20">Secondary demo &amp; test</td>
    </tr>
    <tr>
      <td width="26" class="profile_light" height="20">7</td>
      <td width="70" bgcolor="#FDC677" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="109" class="profile_light" height="20">
      <a href="http://64.147.13.101/tsweb">Search 7</a></td>
      <td width="83" class="profile_light" height="20">0</td>
      <td width="98" class="profile_light" height="20">3</td>
      <td width="76" class="profile_light" height="20">0</td>
      <td width="105" class="profile_light" height="20">1000</td>
      <td width="127" class="profile_light" height="20">N/A</td>
      <td width="94" class="profile_light" height="20">N/A</td>
      <td width="97" class="profile_light" height="20">&nbsp;</td>
      <td width="126" class="profile_light" height="20">off-line</td>
      <td width="635" class="profile_light" height="20">&nbsp;</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" cellspacing="0" height="4">
    <tr>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
      <td background="images/ruler.gif"></td>
    </tr>
  </table>
  <p>&nbsp;|
  <a href="http://orion.mysymmetry.net/CARS/enable_alert.asp">Enable</a> |
  <a href="http://orion.mysymmetry.net/CARS/disable_alert.asp">Disable</a> | 
  Purge | Add Note | </p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>