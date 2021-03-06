<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. |  Alerts</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="javascript" src="inc/string.js"></script>
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


function isName(str)
{
	    return true;    
}

function AtTrim(s)
{
	var r1, r2, s1, s2, s3;

    r1 = new RegExp("^ *");
    r2 = new RegExp(" *$");

    s1 = ""+s+"";
    s2 = s1.replace(r1, "");
    s3 = s2.replace(r2, "");
    
    r1 = null;
    r2 = null;

    return(s3);
}


function CreateAlert()
{ 
	
	document.create_alert.company1.value = AtTrim(document.create_alert.company1.value);
	document.create_alert.company2.value = AtTrim(document.create_alert.company2.value);
	if(document.create_alert.company1.value=="")
	{
		alert("Please input the first company name.");  
		document.create_alert.company1.focus();
		return false ;
	} 
	
	if (!isName(document.create_alert.car_type1.value))
    	{	alert("The name can't include the '<' or '>'")
		document.create_alert.car_type1.focus();
		return false ;
        }
        
	if((document.create_alert.company1.value=="ANY") && (document.create_alert.company1.value==document.create_alert.company2.value))
			{
			alert("You may not use ANY for both companies. Please input valid company.");  
			document.create_alert.company2.focus();
			return false ;
			}
             
	document.create_alert.submit();
	return true;
			
}

function CheckCompanies()
{
  
	if((document.create_alert.company1.value=="ANY") && (document.create_alert.company1.value==document.create_alert.company2.value))
			{
			alert("You may not use ANY for both companies. Please input a valid company in one or both.");  
			document.create_alert.company2.focus();
			return true ;
			}


}


function CheckCarTypes()
{
  
	if((document.create_alert.car_type1.value=="ANY") && (document.create_alert.car_type1.value==document.create_alert.car_type2.value))
			{
			alert("You may not use ANY for both car types. Please input a valid car type in one or both.");  
			document.create_alert.car_type2.focus();
			return true ;
			}


}



function Eatdirt()
{
	if (!isValidEmail(document.create_alert.V_Email.value))
	{
		alert("The email address is invalid. Please input again.");  
		document.create_alert.V_Email.focus();
		return false ;

	}
	//document.create_alert.v_Country_Code.value=parseDigits(document.create_alert.v_Country_Code.value);
	//document.create_alert.v_Area_Code.value=parseDigits(document.create_alert.v_Area_Code.value);
	//document.create_alert.v_Phone_Code.value=parseDigits(document.create_alert.v_Phone_Code.value);
	//document.create_alert.v_Exter_Code.value=parseDigits(document.create_alert.v_Exter_Code.value);
	 if (isNaN(document.create_alert.v_Country_Code.value))	 
	 {  
		alert("The phone country code should be a number. Please input again.");        
		document.create_alert.v_Country_Code.focus();
		return false ;
	 }
	 if (isNaN(document.create_alert.v_Area_Code.value))	 
	 {  
		alert("The phone area or city code can contain only numbers. Please try again.");        
		document.create_alert.v_Area_Code.focus();
		return false ;
	 }
	 if (isNaN(document.create_alert.v_Phone_Code.value))	 
	 {  
		alert("The phone number should be a number. Please input again.");        
		document.create_alert.v_Phone_Code.focus();
		return false ;
	 }
	 if (isNaN(document.create_alert.v_Exter_Code.value))	 
	 {  
		alert("The phone extension should be a number. Please input again.");        
		document.create_alert.v_Exter_Code.focus();
		return false ;
	 }
	             
	document.create_alert.submit();
	return true;
}
//
-->
</script>
<style>
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
-->
</style>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload=" MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <!-- #INCLUDE FILE="inc/page_header_buttons.htm" -->
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/med_bar_tile.gif">
    <img src="images/med_bar.gif" width="12" height="8"></td>
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
                <td align="right">
                <div align="right">
                  <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div>
                </td>
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
    <td background="images/h_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0" id="table1">
      <tr>
        <td><img src="images/h_alerts.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;
       
<!-- JUSTTABS TOP OPEN -->
<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="1132" bgcolor="#FFFFFF">
<tr height="1">
<td colspan="1" width="10">&nbsp;</td>
<td rowspan="2" width="388"><a href="javascript:not_enabled()"><img src="images/loginalerts0_a.JPG" width="92" height="25" hspace="0" vspace="0" border="0" alt="Login Alerts Maint." description=""></a><a href="alerts_notification_car.asp"><img src="images/notificationalerts1_ia.JPG" width="127" height="25" hspace="0" vspace="0" border="0" alt="Notitification Alerts" description=""></a><a href="alerts_rate_management_car.asp"><img src="images/ratemanagementalerts2_ia.JPG" width="169" height="25" hspace="0" vspace="0" border="0" alt="Rate Management" description=""></a></td>
<td colspan="1" >&nbsp;</td>
</tr>
<tr height="1">
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
<td bgcolor="#000000" colspan="1" height="1"><img src=pixel.gif width="1" height="1"></td>
</tr>
</table>
</div>
<table cellpadding="0" cellspacing="0" border="0" ALIGN="CENTER" width="100%" bgcolor="#FFFFFF">
<tr >
<td  width="1" bgcolor="#000000"><img src=pixel.gif width="1" height="1"></td>
<td colspan=3 bgcolor="#D9DEE1">
<table border="0" cellspacing="5" cellpadding="5">
<tr><td>
<font color="#080000">
<P>
<!-- JUSTTABS TOP END TOP -->
&nbsp;
       
              <form method="POST" action="" name="login_alerts" class="search">
</form>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="100%" cellspacing="0" height="4">
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
<table width="745" border="0" cellspacing="0" cellpadding="0" background="images/alt_color.gif">
  <tr>
    <td>
    <table width="1110" border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" background="images/alt_color.gif">
      <tr valign="bottom">
        <td width="10" height="51">&nbsp;</td>
        <td width="179" valign="middle" height="51"><font face="Verdana">Login 
        Alerts!</font></td>
        <td width="583" colspan="3" height="51">
        <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Listed below 
        is a list of the login alerts that have occurred since your last login.</font></td>
        <td width="336" height="51">&nbsp;</td>
      </tr>
      <tr valign="bottom">
        <td width="10" height="22">&nbsp;</td>
        <td width="179" height="22">&nbsp;</td>
        <td width="177" height="22">&nbsp;</td>
        <td width="80" height="22">&nbsp;</td>
        <td width="662" colspan="2" height="22">&nbsp;</td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<table width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td width="169">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">|&lt;</a>
    <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&lt;</a> Page 1 
    of 1 <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&gt;</a>
    <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&gt;|</a></font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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
<table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" id="profiles">
  <tr>
    <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="30">&nbsp;</td>
    <td class="profile_header" width="66" style="background-color: #E07D1A" height="45">
    Selected</td>
    <td class="profile_header" width="65" height="45">Created By</td>
    <td class="profile_header" width="340" height="45">Situation</td>
    <td class="profile_header" width="261" height="45">Quantity &amp; Period</td>
    <td class="profile_header" width="48" height="45">Type</td>
    <td class="profile_header" width="72" height="45">Locations</td>
    <td class="profile_header" width="297" height="45">Recipients</td>
    <td class="profile_header" width="126" height="45">Search Type / Profile</td>
  </tr>
  <tr>
    <td width="26" class="profile_light" height="20">&nbsp;</td>
    <td width="70" bgcolor="#FDC677" align="center" height="20">
    <input type="radio" value="V1" name="selected"></td>
    <td width="65" class="profile_light" height="20">&nbsp;</td>
    <td width="340" class="profile_light" height="20">&nbsp;</td>
    <td width="261" class="profile_light" height="20">&nbsp;</td>
    <td width="48" class="profile_light" height="20">&nbsp;</td>
    <td width="72" class="profile_light" height="20">&nbsp;</td>
    <td width="297" class="profile_light" height="20">&nbsp;</td>
    <td width="126" class="profile_light" height="20">&nbsp;</td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" cellspacing="0" height="4">
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


<p>&nbsp;| <a href="http://orion.mysymmetry.net/CARS/delete_alert.asp">Delete</a> 
| <a href="http://orion.mysymmetry.net/CARS/disable_alert.asp">Disable</a> |</p>
<!-- Content goes before this comment -->
<!-- JUSTTABS BOTTOM OPEN -->
</font></td></tr></table>
</td>
<td  width="1" bgcolor="#000000"><img src=pixel.gif width="1" height="1"></td>
</tr>
<tr bgcolor="#000000" height="1">
<td colspan=5><img src=pixel.gif width="1" height="1"></td>
</tr>
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
