<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" --> 

<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Alerts</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language=javascript src="inc/string.js"></script>

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
//-->
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

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">
    <table width="400" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/b_left.jpg" width="62" height="32"></td>
        <td>
        <a href="search_profiles_car.asp" onMouseOver="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td>
        <td>
        <a href="search_queue_car.asp" onMouseOver="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
        <td>
        <a href="search_criteria_car.asp" onMouseOver="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onMouseOver="MM_swapImage('ra','','images/b_rate_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td>
        <td>
        <a href="searched_alerts.asp" onMouseOver="MM_swapImage('al','','images/b_alert_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onMouseOver="MM_swapImage('us','','images/b_user_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onMouseOver="MM_swapImage('sy','','images/b_system_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td>
      </tr>
    </table>
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
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>&nbsp;</td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<form method="POST" action="searched_alerts.asp" name="search_alerts">
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
      <table width="1108" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="153" background="images/alt_color.gif">
        <tr valign="bottom">
          <td width="10" height="51">&nbsp;</td>
          <td width="179" valign="middle" height="51">
          <img border="0" src="images/search.GIF"></td>
          <td width="583" colspan="3" height="51">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">To search 
          for an Alert, enter the address of the recipient, or a portion of the 
          address. You may also enter the alert type.</font></td>
          <td width="336" height="51">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="26">&nbsp;</td>
          <td width="179" height="26">&nbsp;</td>
          <td width="177" height="26">
          <font face="Verdana, Arial, Helvetica, sans-serif" size="2">
          Recipient's address:</font> </td>
          <td width="80" height="26">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input type="text" name="name" size="20" style="width:150" style="width:150"></font></td>
          <td width="662" colspan="2" height="26">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input name="search" type="submit" id="Open2224" value="    Search    " class="rh_button"></font></td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="22">&nbsp;</td>
          <td width="179" height="22">&nbsp;</td>
          <td width="177" height="22">
          <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Alert 
          Type:</font></td>
          <td width="80" height="22">
          <select size="1" name="type" style="width:150">
          <option selected value="0">Any type</option>
          <option value="1">Email notification</option>
          <option value="2">Pager notification</option>
          <option value="3">Login list</option>
          </select></td>
          <td width="662" colspan="2" height="22">&nbsp;</td>
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
      <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&lt;</a> Page 
      1 of 1 <a href="http://orion.mysymmetry.net/CARS/search_profiles.asp">&gt;</a>
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
      <td width="26" class="profile_light" height="20">1e</td>
      <td width="70" bgcolor="#FDC677" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="65" class="profile_light" height="20">mmeyer</td>
      <td width="340" class="profile_light" height="20">Hertz ICAR <b>&gt;</b> Avis 
      ICAR by 10%</td>
      <td width="261" class="profile_light" height="20">1 event(s) in 24 hour(s)</td>
      <td width="48" class="profile_light" height="20">email</td>
      <td width="72" class="profile_light" height="20">SNA, LAX</td>
      <td width="297" class="profile_light" height="20">michaelm@rate-city.com</td>
      <td width="126" class="profile_light" height="20">As searched</td>
    </tr>
    <tr>
      <td width="26" class="profile_dark" height="20">2d</td>
      <td width="70" bgcolor="#E07D1A" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="65" class="profile_dark" height="20">mmeyer</td>
      <td width="340" class="profile_dark" height="20">Avis ANY <b>&lt;</b> Hertz ANY 
      by 5.00 </td>
      <td width="261" class="profile_dark" height="20">1 event(s) in 24 hour(s)</td>
      <td width="48" class="profile_dark" height="20">Login</td>
      <td width="72" class="profile_dark" height="20">ANY</td>
      <td width="297" class="profile_dark" height="20">michaelm@rate-city.com</td>
      <td width="126" class="profile_dark" height="20">LAX - Hertz 3</td>
    </tr>
    <tr>
      <td width="26" class="profile_light" height="20">3e</td>
      <td width="70" bgcolor="#FDC677" align="center" height="20">
      <input type="radio" value="V1" name="selected"></td>
      <td width="65" class="profile_light" height="20">mmeyer</td>
      <td width="340" class="profile_light" height="20">Hertz ECAR <b>&lt;=</b> 
      60.00</td>
      <td width="261" class="profile_light" height="20">5 event(s) in 120 hour(s)</td>
      <td width="48" class="profile_light" height="20">email</td>
      <td width="72" class="profile_light" height="20">FCA</td>
      <td width="297" class="profile_light" height="20">michaelm@rate-city.com</td>
      <td width="126" class="profile_light" height="20">As searched</td>
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
  | <a href="http://orion.mysymmetry.net/CARS/copy_alert.asp">Copy</a> |
  <a href="http://orion.mysymmetry.net/CARS/enable_alert.asp">Enable</a> |
  <a href="http://orion.mysymmetry.net/CARS/disable_alert.asp">Disable</a> |</p>
  <input type="hidden" name="refresh_from" value="search">
</form>
<form name="create_alert" method="POST" action="searched_alerts.asp" OnSubmit="CreateAlert();return false">
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
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="new_alert" background="images/alt_color.gif" height="561">
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25"><img border="0" src="images/maintenance.GIF"></td>
      <td width="200" height="25"><font face="Verdana" size="2">1. Company</font></td>
      <td width="672" height="25">
      <input type="text" name="company1" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(use &quot;any&quot; to 
      denote any company)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="23">&nbsp;</td>
      <td width="247" height="23">&nbsp;</td>
      <td width="200" height="23"><font face="Verdana" size="2">2. Car</font>
      <font face="Verdana" size="2">Type</font></td>
      <td width="672" height="23">
      <input type="text" name="car_type1" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
      <td width="169" height="23">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(use &quot;any&quot; to 
      denote any car type)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">3. Situation</font></td>
      <td width="672" height="22">
      <select size="1" name="situation" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="0">(None selected)</option>
      <option value="1">&gt; (Greater than)</option>
      <option value="2">&gt;= (Greater than or Equal to)</option>
      <option value="3">= (Equal to)</option>
      <option value="4">&lt;= (Less than or Equal to)</option>
      <option value="5">&lt; (Less than)</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="23">&nbsp;</td>
      <td width="247" height="23">&nbsp;</td>
      <td width="200" height="23"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      Amount</font></td>
      <td width="672" height="23">
      <input type="text" name="situation_amount" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
      <td width="169" height="23">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(enter either 
      monetary units or percent, use XX% formatting for percent) </font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">4. Company</font></td>
      <td width="672" height="22">
      <input type="text" name="company2" size="20" style="width:200; font-family:Verdana; font-size:10pt" OnBlur="CheckCompanies();return false"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(use &quot;any&quot; to 
      denote any company, you may not use &quot;any&quot; for both items 1 &amp; 4)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">5. Car Type</font></td>
      <td width="672" height="22">
      <input type="text" name="car_type2" size="20" style="width:200; font-family:Verdana; font-size:10pt" OnBlur="CheckCarTypes();return false"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(use &quot;any&quot; to 
      denote any car type, you may not use &quot;any&quot; for both items 2 &amp; 5)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">6. Quantity &amp; 
      Period:</font></td>
      <td width="672" height="22">
      <select size="1" name="period" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">After X Events</option>
      <option value="2">After X Events in X Hours</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      No. of Events</font></td>
      <td width="672" height="22">
      <input type="text" name="period_events" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      No. of Hours</font></td>
      <td width="672" height="22">
      <input type="text" name="period_hours" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">7. Alert Type</font></td>
      <td width="672" height="22">
      <select size="1" name="alert_type" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">Email notification</option>
      <option value="2">Pager notification</option>
      <option value="3">Login list</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">8. Locations</font></td>
      <td width="672" height="22">
      <input type="text" name="locations" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(enter 
      airport/city codes &quot;any&quot; for any location) </font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">9. Recipient</font>:</td>
      <td width="672" height="22">
      <input type="text" name="recipient" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19"><font face="Verdana" size="2">(separate each 
      recipient with a comma)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">10. Search Type</font></td>
      <td width="672" height="22">
      <select size="1" name="search_type" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">As searched (all searches)</option>
      <option value="2">Link to Profile</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="61">&nbsp;</td>
      <td width="247" height="61">&nbsp;</td>
      <td width="872" colspan="2" height="61">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<p>
          <input name="submit_alert" type="submit" id="submit_alert" value="    Create   " class="rh_button"></font></td>
      <td width="169" height="61">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="872" colspan="2" height="19">&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
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
  <input type="hidden" name="refresh_from" value="create">
</form>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<script language="javascript">
	document.search_alerts.name.focus();
</script>