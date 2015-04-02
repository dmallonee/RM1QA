<%@ Language=VBScript %>


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
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
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
        <a href="alerts_maint_new.asp" onMouseOver="MM_swapImage('al','','images/b_alert_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onMouseOver="MM_swapImage('us','','images/b_user_on.gif',1)" onMouseOut="MM_swapImgRestore()">
        <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
        <td>
        <a onMouseOver="MM_swapImage('sy','','images/b_system_on.gif',1)" onMouseOut="MM_swapImgRestore()" href="javascript:not_enabled()">
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
<form method="POST" action="searched_alerts.asp" name="search_alerts" class="search">
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
          <input type="text" name="name" size="20" style="width:150" style="width:150" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');"></font></td>
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
          <select size="1" name="type" style="border:1px solid #000000; width:150; background-color:#FF9933">
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
  </form>
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
      <td width="26" class="profile_light" height="20">&nbsp;</td>
      <td width="70" bgcolor="#FDC677" align="center" height="20">
      <input type="radio" value="V1" name="selected" class="nothing"></td>
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
  | <a href="http://orion.mysymmetry.net/CARS/copy_alert.asp">Copy</a> |
  <a href="http://orion.mysymmetry.net/CARS/enable_alert.asp">Enable</a> |
  <a href="http://orion.mysymmetry.net/CARS/disable_alert.asp">Disable</a> |</p>
  <input type="hidden" name="refresh_from" value="search">
</form>
<form name="create_alert" method="POST" action="alerts_maint_new.asp" OnSubmit="CreateAlert();return false">
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
      <td width="247" height="25">
      <img border="0" src="images/maintenance.GIF" width="162" height="25"></td>
      <td width="200" height="25"><font face="Verdana" size="2">1. Rate Change 
      Alert No.:</font></td>
      <td width="672" height="25">
      <input type="text" name="rate_alert_number" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right; background-image:url('images/alt_color.gif')" value="231" READONLY></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25">&nbsp;</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">2. Description:</font></td>
      <td width="672" height="25">
      <input type="text" name="description" size="20" style="width:439; font-family:Verdana; font-size:10pt; height:21"></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25">&nbsp;</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">3. Rule Begin 
      Date:</font></td>
      <td width="672" height="25">
                <input type="text" name="begin_date" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')" value="continuous">
                <a href="javascript:cal6.popup();">
                <img src="images/cal.gif" width="16" height="16" border="0" alt="Click here to select the start date"></a></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25">&nbsp;</td>
      <td width="672" height="25">
      <font face="Verdana" size="2">(enter 'continuous' or blank for no begin date) </font></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">4. Rule End 
      Date:</font></td>
      <td width="672" height="25">
                <input type="text" name="begin_date0" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')" value="continuous">
                <a href="javascript:cal6.popup();">
                <img src="images/cal.gif" width="16" height="16" border="0" alt="Click here to select the start date"></a></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25">&nbsp;</td>
      <td width="672" height="25">
      <font face="Verdana" size="2">(enter 'continuous' or blank for no end date)</font></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">5. First 
      Pick-up:</font></td>
      <td width="672" height="25">
                <input type="text" name="pick_up_datetime" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')" value="12/12/2004 12:59 PM">
                <a href="javascript:cal6.popup();">
                <img src="images/cal.gif" width="16" height="16" border="0" alt="Click here to select the start date"></a></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25">&nbsp;</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">6. Last Pick-up:</font></td>
      <td width="672" height="25">
                <input type="text" name="drop_off_datetime" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')">
                <a href="javascript:cal6.popup();">
                <img src="images/cal.gif" width="16" height="16" border="0" alt="Click here to select the start date"></a></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25">&nbsp;</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">7. System Rate 
      Code:</font></td>
      <td width="672" height="25">
      <input type="text" name="rate_code" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25">&nbsp;</td>
      <td width="672" height="25"><font face="Verdana" size="2">
      (the rate code used within your system, i.e. Daily, Weekly, etc.)</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">8. Car 
      Companies:</font></td>
      <td width="672" height="25">&nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" style="padding-top: 2px">
      <font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;Competitive Set:</font></td>
      <td width="672" height="25">
      <select size="4" name="company" multiple style="width:200; font-family:Verdana; font-size:10pt" >
      <option>ANY COMPANY</option>
      <option value="ACE">ACE </option>
<option value="ADVANTAGE">ADVANTAGE</option>
<option value="ALAMO">ALAMO </option>
<option value="AVIS">AVIS </option>
<option value="BUDGET">BUDGET </option>
<option value="DOLLAR">DOLLAR </option>
<option value="ENTERPRISE">ENTERPRISE </option>
<option value="FOX">FOX </option>
<option value="HERTZ">HERTZ </option>
<option value="NATIONAL">NATIONAL </option>
<option value="THRIFTY">THRIFTY </option>
      </select></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" valign="top" style="padding-top: 2px">
      &nbsp;</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" style="padding-top: 2px">
      <font face="Verdana" size="2">&nbsp;&nbsp; Comparison Company<br>
&nbsp;&nbsp; (usually self):</font></td>
      <td width="672" height="25">
      <select size="4" name="company5" multiple style="width:200; font-family:Verdana; font-size:10pt; background-image:url('images/alt_color.gif'); background-repeat:repeat" >
      <option value="ACE">ACE </option>
<option value="ADVANTAGE" selected>ADVANTAGE</option>
<option value="ALAMO">ALAMO </option>
<option value="AVIS">AVIS </option>
<option value="BUDGET">BUDGET </option>
<option value="DOLLAR">DOLLAR </option>
<option value="ENTERPRISE">ENTERPRISE </option>
<option value="FOX">FOX </option>
<option value="HERTZ">HERTZ </option>
<option value="NATIONAL">NATIONAL </option>
<option value="THRIFTY">THRIFTY </option>
      </select></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" valign="top" style="padding-top: 2px">
      &nbsp;</td>
      <td width="672" height="25">
      <font face="Verdana" size="2">(use &quot;any&quot; to 
      denote any company, you may not use &quot;any&quot; for <br>
      both items)</font></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" valign="top" style="padding-top: 2px">
      <font face="Verdana" size="2">9. LOR(s):</font></td>
      <td width="672" height="25">
      <select size="4" name="company3" multiple style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">1</option>
      <option value="5">5</option>
      </select></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" valign="top" style="padding-top: 2px">&nbsp;</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="71">&nbsp;</td>
      <td width="247" height="71">
      &nbsp;</td>
      <td width="200" height="71" valign="top" style="padding-top: 2px">
      <font face="Verdana" size="2">10. Location(s):</font></td>
      <td width="672" height="71" valign="top">
      <select size="4" name="company4" multiple style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="LAX">LAX</option>
      <option value="LAX-RES01">LAX-RES01</option>
      <option value="SAN">SAN</option>
      <option value="SFO">SFO</option>
      <option value="SNA">SNA</option>
      </select> <button name="custom_location" style="height: 24">Edit</button></td>
      <td width="169" height="71">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="24">&nbsp;</td>
      <td width="247" height="24">
      &nbsp;</td>
      <td width="200" height="24" valign="top" style="padding-top: 2px">&nbsp;</td>
      <td width="672" height="24">
      &nbsp;</td>
      <td width="169" height="24">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" style="vertical-align: top; padding-top: 2px">
      <font face="Verdana" size="2">11. Car Types:</font></td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" style="padding-top: 2px">
      <font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Car Type(s):</font></td>
      <td width="672" height="25">
      <p style="margin-top: 2px">
      <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
      <select name="selected_car_types" size="4" multiple style="width:200; font-family:Verdana; font-size:10pt">
      <option value="ANYY">ANY TYPE</option>
      <option value="CCAR">CCAR</option>
      <option value="CCMN">CCMN</option>
      <option value="CCMR">CCMR</option>
      <option value="CDAN">CDAN</option>
      <option value="CDAR">CDAR</option>
      <option value="CDMN">CDMN</option>
      <option value="CDMR">CDMR</option>
      <option value="CFAR">CFAR</option>
      <option value="CPAR">CPAR</option>
      <option value="CVMR">CVMR</option>
      <option value="CWAN">CWAN</option>
      <option value="CWAR">CWAR</option>
      <option value="CWMN">CWMN</option>
      <option value="CWMR">CWMR</option>
      <option value="CXMN">CXMN</option>
      <option value="EBMN">EBMN</option>
      <option value="ECAN">ECAN</option>
      <option value="ECAR">ECAR</option>
      <option value="ECMN">ECMN</option>
      <option value="ECMR">ECMR</option>
      <option value="EDAN">EDAN</option>
      <option value="EDAR">EDAR</option>
      <option value="EDMN">EDMN</option>
      <option value="EDMR">EDMR</option>
      <option value="FCAN">FCAN</option>
      <option value="FCAR">FCAR</option>
      <option value="FCMN">FCMN</option>
      <option value="FCMR">FCMR</option>
      <option value="FDAN">FDAN</option>
      <option value="FDAR">FDAR</option>
      <option value="FDMN">FDMN</option>
      <option value="FDMR">FDMR</option>
      <option value="FFAR">FFAR</option>
      <option value="FPAR">FPAR</option>
      <option value="FVAN">FVAN</option>
      <option value="FVAR">FVAR</option>
      <option value="FVMN">FVMN</option>
      <option value="FVMR">FVMR</option>
      <option value="FWAR">FWAR</option>
      <option value="FWMN">FWMN</option>
      <option value="FWMR">FWMR</option>
      <option value="ICAN">ICAN</option>
      <option value="ICAR">ICAR</option>
      <option value="ICMN">ICMN</option>
      <option value="ICMR">ICMR</option>
      <option value="IDAN">IDAN</option>
      <option value="IDAR">IDAR</option>
      <option value="IDMN">IDMN</option>
      <option value="IDMR">IDMR</option>
      <option value="IFAR">IFAR</option>
      <option value="IJAR">IJAR</option>
      <option value="IPAR">IPAR</option>
      <option value="IVMN">IVMN</option>
      <option value="IVMR">IVMR</option>
      <option value="IWAN">IWAN</option>
      <option value="IWAR">IWAR</option>
      <option value="IWMN">IWMN</option>
      <option value="IWMR">IWMR</option>
      <option value="IXMN">IXMN</option>
      <option value="IXMR">IXMR</option>
      <option value="LCAR">LCAR</option>
      <option value="LDAR">LDAR</option>
      <option value="LDMR">LDMR</option>
      <option value="LFAR">LFAR</option>
      <option value="LTAR">LTAR</option>
      <option value="LWAR">LWAR</option>
      <option value="LXAR">LXAR</option>
      <option value="MBMN">MBMN</option>
      <option value="MCAR">MCAR</option>
      <option value="MCMN">MCMN</option>
      <option value="MCMR">MCMR</option>
      <option value="MVAR">MVAR</option>
      <option value="PCAR">PCAR</option>
      <option value="PCMR">PCMR</option>
      <option value="PDAR">PDAR</option>
      <option value="PDMR">PDMR</option>
      <option value="PFAR">PFAR</option>
      <option value="PSAR">PSAR</option>
      <option value="PVAN">PVAN</option>
      <option value="PVMN">PVMN</option>
      <option value="PVMR">PVMR</option>
      <option value="PWAR">PWAR</option>
      <option value="PXAR">PXAR</option>
      <option value="SCAN">SCAN</option>
      <option value="SCAR">SCAR</option>
      <option value="SCMN">SCMN</option>
      <option value="SCMR">SCMR</option>
      <option value="SDAN">SDAN</option>
      <option value="SDAR">SDAR</option>
      <option value="SDMN">SDMN</option>
      <option value="SDMR">SDMR</option>
      <option value="SFAR">SFAR</option>
      <option value="SJAR">SJAR</option>
      <option value="SPAR">SPAR</option>
      <option value="SSAR">SSAR</option>
      <option value="STAR">STAR</option>
      <option value="SVAR">SVAR</option>
      <option value="SVMN">SVMN</option>
      <option value="SVMR">SVMR</option>
      <option value="SWAN">SWAN</option>
      <option value="SWAR">SWAR</option>
      <option value="SWMN">SWMN</option>
      <option value="SWMR">SWMR</option>
      <option value="XCAR">XCAR</option>
      <option value="XCMN">XCMN</option>
      <option value="XDAR">XDAR</option>
      <option value="XDMN">XDMN</option>
      <option value="XFAR">XFAR</option>
      <option value="XRAR">XRAR</option>
      <option value="XSAR">XSAR</option>
      <option value="XXAR">XXAR</option>
      </select></font></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25" style="vertical-align: top; padding-top: 2px">
      &nbsp;</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Comparison Car&nbsp;&nbsp;&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp; Type(s):</font></td>
      <td width="672" height="19">
      <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
      <select name="selected_car_types0" size="4" multiple style="width:200; font-family:Verdana; font-size:10pt">
      <option value="ANYY">ANY TYPE</option>
      <option value="CCAR">CCAR</option>
      <option value="CCMN">CCMN</option>
      <option value="CCMR">CCMR</option>
      <option value="CDAN">CDAN</option>
      <option value="CDAR">CDAR</option>
      <option value="CDMN">CDMN</option>
      <option value="CDMR">CDMR</option>
      <option value="CFAR">CFAR</option>
      <option value="CPAR">CPAR</option>
      <option value="CVMR">CVMR</option>
      <option value="CWAN">CWAN</option>
      <option value="CWAR">CWAR</option>
      <option value="CWMN">CWMN</option>
      <option value="CWMR">CWMR</option>
      <option value="CXMN">CXMN</option>
      <option value="EBMN">EBMN</option>
      <option value="ECAN">ECAN</option>
      <option value="ECAR">ECAR</option>
      <option value="ECMN">ECMN</option>
      <option value="ECMR">ECMR</option>
      <option value="EDAN">EDAN</option>
      <option value="EDAR">EDAR</option>
      <option value="EDMN">EDMN</option>
      <option value="EDMR">EDMR</option>
      <option value="FCAN">FCAN</option>
      <option value="FCAR">FCAR</option>
      <option value="FCMN">FCMN</option>
      <option value="FCMR">FCMR</option>
      <option value="FDAN">FDAN</option>
      <option value="FDAR">FDAR</option>
      <option value="FDMN">FDMN</option>
      <option value="FDMR">FDMR</option>
      <option value="FFAR">FFAR</option>
      <option value="FPAR">FPAR</option>
      <option value="FVAN">FVAN</option>
      <option value="FVAR">FVAR</option>
      <option value="FVMN">FVMN</option>
      <option value="FVMR">FVMR</option>
      <option value="FWAR">FWAR</option>
      <option value="FWMN">FWMN</option>
      <option value="FWMR">FWMR</option>
      <option value="ICAN">ICAN</option>
      <option value="ICAR">ICAR</option>
      <option value="ICMN">ICMN</option>
      <option value="ICMR">ICMR</option>
      <option value="IDAN">IDAN</option>
      <option value="IDAR">IDAR</option>
      <option value="IDMN">IDMN</option>
      <option value="IDMR">IDMR</option>
      <option value="IFAR">IFAR</option>
      <option value="IJAR">IJAR</option>
      <option value="IPAR">IPAR</option>
      <option value="IVMN">IVMN</option>
      <option value="IVMR">IVMR</option>
      <option value="IWAN">IWAN</option>
      <option value="IWAR">IWAR</option>
      <option value="IWMN">IWMN</option>
      <option value="IWMR">IWMR</option>
      <option value="IXMN">IXMN</option>
      <option value="IXMR">IXMR</option>
      <option value="LCAR">LCAR</option>
      <option value="LDAR">LDAR</option>
      <option value="LDMR">LDMR</option>
      <option value="LFAR">LFAR</option>
      <option value="LTAR">LTAR</option>
      <option value="LWAR">LWAR</option>
      <option value="LXAR">LXAR</option>
      <option value="MBMN">MBMN</option>
      <option value="MCAR">MCAR</option>
      <option value="MCMN">MCMN</option>
      <option value="MCMR">MCMR</option>
      <option value="MVAR">MVAR</option>
      <option value="PCAR">PCAR</option>
      <option value="PCMR">PCMR</option>
      <option value="PDAR">PDAR</option>
      <option value="PDMR">PDMR</option>
      <option value="PFAR">PFAR</option>
      <option value="PSAR">PSAR</option>
      <option value="PVAN">PVAN</option>
      <option value="PVMN">PVMN</option>
      <option value="PVMR">PVMR</option>
      <option value="PWAR">PWAR</option>
      <option value="PXAR">PXAR</option>
      <option value="SCAN">SCAN</option>
      <option value="SCAR">SCAR</option>
      <option value="SCMN">SCMN</option>
      <option value="SCMR">SCMR</option>
      <option value="SDAN">SDAN</option>
      <option value="SDAR">SDAR</option>
      <option value="SDMN">SDMN</option>
      <option value="SDMR">SDMR</option>
      <option value="SFAR">SFAR</option>
      <option value="SJAR">SJAR</option>
      <option value="SPAR">SPAR</option>
      <option value="SSAR">SSAR</option>
      <option value="STAR">STAR</option>
      <option value="SVAR">SVAR</option>
      <option value="SVMN">SVMN</option>
      <option value="SVMR">SVMR</option>
      <option value="SWAN">SWAN</option>
      <option value="SWAR">SWAR</option>
      <option value="SWMN">SWMN</option>
      <option value="SWMR">SWMR</option>
      <option value="XCAR">XCAR</option>
      <option value="XCMN">XCMN</option>
      <option value="XDAR">XDAR</option>
      <option value="XDMN">XDMN</option>
      <option value="XFAR">XFAR</option>
      <option value="XRAR">XRAR</option>
      <option value="XSAR">XSAR</option>
      <option value="XXAR">XXAR</option>
      </select></font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">
      <font face="Verdana" size="2">(use &quot;any&quot; to 
      denote any car type, you may not use &quot;any&quot; for<br>
&nbsp;both items)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19"><font face="Verdana" size="2">12. Rate Source::</td>
      <td width="672" height="19">
      <select size="1" name="location" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="CAR">Brand sites (avis.com, etc.)</option>
      <option value="ORB">Orbitz</option>
      <option value="EXP">Expedia</option>
      <option value="TRV">Travelocity</option>
      <option value="ATV">All Travel sites</option>
      <option value="ALL">All sites</option>
      </select>&nbsp;</td>
      <td width="169" height="19">&nbsp;</td>
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
      <td width="200" height="22"><font face="Verdana" size="2">13. Situation:</font></td>
      <td width="672" height="22">
      <select size="1" name="situation" style="width:200; font-family:Verdana; font-size:10pt; height:24">
      <option selected value="0">(None selected)</option>
      <option value="1">> (all competitors)</option>
      <option value="2">> (any competitor)</option>
      <option value="3">> (custom)</option>
      <option value="4">&gt;= (all competitors)</option>
      <option value="5">>= (any competitor)</option>
      <option value="6">>= (custom)</option>
      <option value="7">= (all competitors)</option>
      <option value="8">= (any competitor)</option>
      <option value="9">= (custom)</option>
      <option value="11">< = (all competitors)</option>
      <option value="12"><= (any competitor)</option>
      <option value="13">&lt;= (custom)</option>
      <option value="14">< (all competitors)</option>
      <option>< (any competitor)</option>
      <option value="15">< (custom)</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="23">&nbsp;</td>
      <td width="247" height="23">&nbsp;</td>
      <td width="200" height="23"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      Amount:</font></td>
      <td width="672" height="23">
      <input type="text" name="situation_amount" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      <br>
      <input type="radio" value="dollar" name="amount_type" style="font-family:Verdana; font-size:10pt" checked><font face="Verdana" size="2">Whole Dollar<br>
      <input type="radio" value="percent" name="amount_type" style="font-family:Verdana; font-size:10pt"><font face="Verdana" size="2">Percentage</td>
      <td width="169" height="23">&nbsp;</td>
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
      <td width="200" height="22" valign="top">&nbsp;</td>
      <td width="672" height="22">
      &nbsp;<td width="169" height="22">&nbsp;</td>
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
      No. of Events:</font></td>
      <td width="672" height="22">
      <input type="text" name="period_events" size="20" style="width:200; font-family:Verdana; font-size:10pt"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      No. of Hours:</font></td>
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
      <td width="200" height="22"><font face="Verdana" size="2">7. Alert Type:</font></td>
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
      <td width="200" height="22"><font face="Verdana" size="2">10. Search Type:</font></td>
      <td width="672" height="22">
      <select size="1" name="search_type" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">As searched (all searches)</option>
      <option value="2">Link to Profile</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
      Profile:</font></td>
      <td width="672" height="22">
      <select size="1" name="profile" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">(none selected)</option>
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">&nbsp;</td>
      <td width="672" height="22">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input name="submit_alert1" type="submit" id="submit_alert0" value="    Create   " class="rh_button"></font></td>
      <td width="672" height="22">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">&nbsp;</td>
      <td width="672" height="22">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">
      <img border="0" src="images/inventory.GIF"></td>
      <td width="200" height="25"><font face="Verdana" size="2">
      <input type="checkbox" name="link_to_arms" value="yes" checked>Link 
      Alert to ARMS</td>
      <td width="672" height="25">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">&nbsp;</td>
      <td width="672" height="22">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">11. System car 
      type:</font></td>
      <td width="672" height="25">
      <input type="text" name="company" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">&nbsp;</td>
      <td width="672" height="22">
      <font face="Verdana" size="2">(separate each 
      car type with a comma)</font></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">12. System rate 
      codes:</font></td>
      <td width="672" height="25">
      <input type="text" name="company" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">&nbsp;</td>
      <td width="672" height="22">
      <font face="Verdana" size="2">(separate each 
      car type with a comma)</font></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">3. Situation:</font></td>
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
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="23"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      Level:</font></td>
      <td width="672" height="23">
      <input type="text" name="situation_amount" size="20" style="width:200; font-family:Verdana; font-size:10pt">
      </td>
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