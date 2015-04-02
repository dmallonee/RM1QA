<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180


	'On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2
	Dim adoCmd3	
	Dim adoRS3
	Dim adoCmd4
	Dim adoRS4


	Dim adoPrices
	Dim strUserId
	Dim intRuleId
	Dim strAlertDesc
	Dim datBeginDate


	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	intRuleId = Request.QueryString("rateruleid")	
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd1 = CreateObject("ADODB.Command")

	adoCmd1.ActiveConnection =  strConn
	adoCmd1.CommandText = "data_source_select"
	adoCmd1.CommandType = 4
		
	Set adoRS1 = adoCmd1.Execute

	Rem Get the vendors
	Set adoCmd2 = CreateObject("ADODB.Command")

	adoCmd2.ActiveConnection =  strConn
	adoCmd2.CommandText = "vendor_select_ex"
	adoCmd2.CommandType = 4
		
	Set adoRS2 = adoCmd2.Execute
	
	Rem Get the vendors
	Set adoCmd3 = CreateObject("ADODB.Command")

	adoCmd3.ActiveConnection =  strConn
	adoCmd3.CommandText = "vendor_select"
	adoCmd3.CommandType = 4
		
	Set adoRS3 = adoCmd3.Execute

	Rem Get the vendors
	Set adoCmd4 = CreateObject("ADODB.Command")

	adoCmd4.ActiveConnection =  strConn
	adoCmd4.CommandText = "car_rate_rule_select"
	adoCmd4.CommandType = 4
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@rate_rule_id", 3, 1, 0, Null)
	adoCmd4.Parameters.Append adoCmd4.CreateParameter("@user_id", 3, 1, 0, strUserId)
			
	Set adoRS4 = adoCmd4.Execute
	Set adoRS18a = adoCmd4.Execute
	Set adoRS18b = adoCmd4.Execute
			
	Rem Get the cities
	Set adoCmd6 = CreateObject("ADODB.Command")

	adoCmd6.ActiveConnection =  strConn
	adoCmd6.CommandText = "user_city_select"
	adoCmd6.CommandType = 4

	adoCmd6.Parameters.Append adoCmd6.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS6 = adoCmd6.Execute

	Rem Get the car types
	Set adoCmd7 = CreateObject("ADODB.Command")

	adoCmd7.ActiveConnection =  strConn
	adoCmd7.CommandText = "car_type_select"
	adoCmd7.CommandType = 4
		
	Set adoRS7 = adoCmd7.Execute
	Set adoRS8 = adoCmd7.Execute
	
	Set adoCmd9 = CreateObject("ADODB.Command")

	adoCmd9.ActiveConnection =  strConn
	adoCmd9.CommandText = "car_shop_profile_select"
	adoCmd9.CommandType = 4

	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@desc", 200, 1, 255)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@vend_cds", 200, 1, 1024)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@user_id", 3, 1, 0, strUserId)
	adoCmd9.Parameters.Append adoCmd9.CreateParameter("@profile_id", 3, 1, 0, Null)
		
	Set adoRS9 = adoCmd9.Execute
	



	If intRuleId > 0 Then

		Rem Get the specific rule
		Set adoCmd5 = CreateObject("ADODB.Command")

		adoCmd5.ActiveConnection = strConn
		adoCmd5.CommandText = "car_rate_rule_select"
		adoCmd5.CommandType = 4
		
		adoCmd5.Parameters.Append adoCmd5.CreateParameter("@rate_rule_id", 3, 1, 0, intRuleId)
	
		Set adoRS5 = adoCmd5.Execute

		intRuleId = adoRS5.Fields("rate_rule_id").Value
		strAlertDesc = adoRS5.Fields("alert_desc").Value
		If IsDate(adoRS5.Fields("begin_dt").Value) Then
			datBeginDate = FormatDateTime(adoRS5.Fields("begin_dt").Value, 2)
		Else
			datBeginDate = "continuous"
		End If
		If IsDate(adoRS5.Fields("end_dt").Value) Then
			datEndDate = FormatDateTime(adoRS5.Fields("end_dt").Value, 2)
		Else
			datEndDate = "continuous"
		End If
		If IsDate(adoRS5.Fields("first_pickup_dt").Value) Then
			datFirstPickupDate = FormatDateTime(adoRS5.Fields("first_pickup_dt").Value, 2)
		Else
			datFirstPickupDate = "continuous"
		End If
		If IsDate(adoRS5.Fields("last_pickup_dt").Value) Then
			datSecondPickupDate = FormatDateTime(adoRS5.Fields("last_pickup_dt").Value, 2)
		Else
			datSecondPickupDate = "continuous"
		End If
		strClientRateCode = adoRS5.Fields("client_sys_rate_cd").Value
		strCityCd = adoRS5.Fields("city_cd").Value
		strVendCd = adoRS5.Fields("vend_cd1").Value
		strSelfCd = adoRS5.Fields("vend_cd2").Value
		intLOR = adoRS5.Fields("lor").Value
		strCarTypeCd1 = adoRS5.Fields("car_type_cd1").Value
		strCarTypeCd2 = adoRS5.Fields("car_type_cd2").Value
		strDataSource = adoRS5.Fields("data_source").Value
		intSituationCd = adoRS5.Fields("situation_cd").Value
		strSituationAmt = adoRS5.Fields("situation_amt").Value
		blnIsDollar = adoRS5.Fields("is_dollar").Value
		strQuantityPeriodCd =  adoRS5.Fields("quantity_period_cd").Value
		intEventCount =  adoRS5.Fields("event_count").Value
		intEventTimeCount =  adoRS5.Fields("event_time_count").Value
		strEventTimeType = adoRS5.Fields("event_time_type").Value
		intResponseCd = adoRS5.Fields("response_cd").Value
		strResponseAmt = adoRS5.Fields("response_amt").Value
		blnResponseDollar = adoRS5.Fields("is_response_dollar").Value
		strRangeMax = adoRS5.Fields("range_max").Value
		strRangeMin = adoRS5.Fields("range_min").Value
		intSuccessId = adoRS5.Fields("on_success_id").Value
		intFailureId = adoRS5.Fields("on_failure_id").Value
		strProfileId = adoRS5.Fields("profile_id").Value
		blnSearchProfile = adoRS5.Fields("search_profile").Value
				
		strButton = "Update"


	Else
	
		intRuleId = "(New - one will be assigned)"
		strAlertDesc = ""
		datBeginDate = "continuous"
		datEndDate = "continuous"
		datFirstPickupDate = "continuous"
		datSecondPickupDate = "continuous"
		strClientRateCode = ""
		strCityCd = ""
		strVendCd = ""
		strSelfCd = ""
		intLOR = "1"
		strCarTypeCd1 = "CCAR"
		strCarTypeCd2 = "CCAR"
		strDataSource = "ORB"
		intSituationCd = 0
		strSituationAmt = ""
		blnIsDollar = 1
		strQuantityPeriodCd  = 0
		intEventCount =  0
		intEventTimeCount =  0
		strEventTimeType = ""
		intResponseCd = 0
		strResponseAmt = 0
		blnResponseDollar = True
		strRangeMax = ""
		strRangeMin = ""
		intSuccessId = 0
		intFailureId = 0
		strProfileId = "XX"
		blnSearchProfile = False
		
		strButton = "Create"

	End If

%>



<!doctype HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; Alerts! | Rate Management</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language='javascript' src="inc/string.js"></script>
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


function CheckResponse()
{

	if (document.create_alert.response_type.selectedIndex != 0)
	{
		document.create_alert.response_amount.disabled = false;
	}
	else if (document.create_alert.response_type.selectedIndex == 0)
	{
		document.create_alert.response_amount.disabled = true;
		document.create_alert.response_amount.value = '';
		
	}

}


function maint_action(action_id)
{
	document.maint.action.value = action_id;
	document.maint.submit();

}

//-->
</script>

<iframe src="calb.htm" style="display:none;position:absolute;width:148;height:194;z-index=100" id="CalFrame" marginheight="0" marginwidth="0" noresize frameborder="0" scrolling="NO">
</iframe>

<script language="JavaScript">

//
// Expedia Style Calendar Control Scripts
//



var cF=document.all.CalFrame;var cW=window.frames.CalFrame;var g_tid=0;var g_cP,g_eD,g_eDP,g_dmin,g_dmax,g_htm;

function CB(){event.cancelBubble=true}

function SCal(cP,eD,eDP,dmin,dmax,htm){clearTimeout(g_tid);var s=(g_eD==eD);g_cP=cP;g_eD=eD;g_eDP=eDP;g_dmin=dmin;g_dmax=dmax;g_htm=htm;WaitCal(true,s);}
function CancelCal(){clearTimeout(g_tid);cF.style.display="none";}
function WaitCal(i,s)
{
	if(null==cW.g_fCL||false==cW.g_fCL)
	{
	if(i)
	{
	if(s&&"block"==cF.style.display){cF.style.display="none";return;}
	
	cW.location.replace(g_htm);
	PosCal(g_cP);
	cF.style.display="block";
	}
	g_tid=setTimeout("WaitCal()", 200);
	}
	else cW.DoCal(g_cP,g_eD,g_eDP,g_dmin,g_dmax);
}

function PosCal(cP)
{
	var dB=document.body;var eL=0;var eT=0;
	for(var p=cP;p&&p.tagName!='BODY';p=p.offsetParent){eL+=p.offsetLeft;eT+=p.offsetTop;}
	var eH=cP.offsetHeight;var dH=cF.style.pixelHeight;var sT=dB.scrollTop;
	if(eT-dH>=sT&&eT+eH+dH>dB.clientHeight+sT)eT-=dH;else eT+=eH;
	cF.style.left=eL;cF.style.top=eT;
}





var cF=document.all.CalFrame;var cW=window.frames.CalFrame;var g_tid=0;var g_cP,g_eD,g_eDP,g_dmin,g_dmax,g_htm;

function CB(){event.cancelBubble=true}
function SCal(cP,eD,eDP,dmin,dmax,htm){clearTimeout(g_tid);var s=(g_eD==eD);g_cP=cP;g_eD=eD;g_eDP=eDP;g_dmin=dmin;g_dmax=dmax;g_htm=htm;WaitCal(true,s);}
function CancelCal(){clearTimeout(g_tid);cF.style.display="none";}
function WaitCal(i,s)
{
	if(null==cW.g_fCL||false==cW.g_fCL)
	{
	if(i)
	{
	if(s&&"block"==cF.style.display){cF.style.display="none";return;}
	
	cW.location.replace(g_htm);
	PosCal(g_cP);
	cF.style.display="block";
	}
	g_tid=setTimeout("WaitCal()", 200);
	}
	else cW.DoCal(g_cP,g_eD,g_eDP,g_dmin,g_dmax);
}

function PosCal(cP)
{
	var dB=document.body;var eL=0;var eT=0;
	for(var p=cP;p&&p.tagName!='BODY';p=p.offsetParent){eL+=p.offsetLeft;eT+=p.offsetTop;}
	var eH=cP.offsetHeight;var dH=cF.style.pixelHeight;var sT=dB.scrollTop;
	if(eT-dH>=sT&&eT+eH+dH>dB.clientHeight+sT)eT-=dH;else eT+=eH;
	cF.style.left=eL;cF.style.top=eT;
}

function GetDowStart() {return 0;}function GetDateFmt() {return "mmddyy";}function GetDateSep() {return "/";}
function ShowCalendar(eP,eD,eDP,dmin,dmax)
{
	var htm="cal.htm";
	SCal(eP,eD,eDP,dmin,dmax,htm);
}

</script>
<script for="document" event="onclick()">
<!--
CancelCal();
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
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<script src="inc/header_menu_support.js" language="javascript"></script>
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

<br>

<!-- JUSTTABS TOP OPEN -->
<div align="center">
<table cellpadding="0" cellspacing="0" border="0" width="1132" bgcolor="#FFFFFF">
<tr height="1">
<td colspan="1" width="10">&nbsp;</td>
<td rowspan="2" width="388"><a href="javascript:not_enabled()">
<img src="images/loginalerts0_ia.JPG" width="92" height="25" hspace="0" vspace="0" border="0" alt="Login Alerts Maint." description=""></a><a href="alerts_notification_car.asp"><img src="images/notificationalerts1_ia.JPG" width="127" height="25" hspace="0" vspace="0" border="0" alt="Notitification Alerts" description=""></a><img src="images/ratemanagementalerts2_a.JPG" width="169" height="25" hspace="0" vspace="0" border="0" alt="Rate Management" description=""></td>
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
<!-- JUSTTABS TOP OPEN-END -->
&nbsp;
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
      <table width="1108" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" background="images/alt_color.gif">
        <tr valign="bottom">
          <td width="10" height="51">&nbsp;</td>
          <td width="179" height="51">
          <img border="0" src="images/search.GIF"></td>
          <td width="583" colspan="3" height="51">
          <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">To search 
          for an Alert, enter a login id, or a portion of the 
          address. You may also enter the alert type.</font></td>
          <td width="336" height="51">&nbsp;</td>
        </tr>
        <tr valign="bottom">
          <td width="10" height="26">&nbsp;</td>
          <td width="179" height="26">&nbsp;</td>
          <td width="177" height="26">
          <font face="Verdana, Arial, Helvetica, sans-serif" size="2">Owner:</font> </td>
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
          <option value="4">Rate Mgmt</option>
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
     <form name="maint" method="POST" action="car_rate_rule_maint.asp">
  <table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" id="profiles">
    <tr>
      <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="30">&nbsp;</td>
      <td class="profile_header" width="58" style="background-color: #E07D1A" height="45">
      <font size="2">Selected</font></td>
      <td class="profile_header" width="291" height="45"><font size="2">
		Description</font></td>
      <td class="profile_header" width="46" height="45"><font size="2">Rate Code</font></td>
      <td class="profile_header" width="94" height="45"><font size="2">Situation</font></td>
      <td class="profile_header" width="63" height="45"><font size="2">Locations</font></td>
      <td class="profile_header" width="153" height="45"><font size="2">
      Recipient or Response</font></td>
      <td class="profile_header" width="349" height="45"><font size="2">Search Type / Profile</font></td>
    </tr>
    
  <%
        
        Dim strClass
        Dim strOrange
        Dim intCount

		While adoRS4.EOF = False
		
			If strClass = "profile_light" Then
				strClass = "profile_dark"
				strOrange = "bgcolor='#E07D1A'"
			Else
				strClass = "profile_light"
				strOrange = "bgcolor='#FDC677'"
			End If
			
			intCount = intCount + 1
			
		%>
  <tr>
    <td class="<%=strClass %>" height="20">
	<%=adoRS4.Fields("rate_rule_id").Value  & LCASE(adoRS4.Fields("rule_status").Value)%></td>
    <td bgcolor="#FDC677" align="center" height="20">
    <input type="checkbox" value="<%=adoRS4.Fields("rate_rule_id").Value %>" name="rate_rule_id"></td>
    <td class="<%=strClass %>" height="20" width="291">
    <a target="_self" title="<%=adoRS4.Fields("alert_desc").Value %>" href="alerts_rate_management_car.asp?rateruleid=<%=adoRS4.Fields("rate_rule_id").Value %>">
    <%=adoRS4.Fields("alert_desc").Value %></a></td>
    <td class="<%=strClass %>" height="20" width="46">
    <font color="#080000">
	<%=adoRS4.Fields("client_sys_rate_cd").Value %></font></td>
    <td class="<%=strClass %>" height="20" width="94">
	<%=adoRS4.Fields("situation_cd").Value %></td>
    <td class="<%=strClass %>" height="20" width="63">
	<% If adoRS4.Fields("city_cd").Value = "" Then %>
	 Any
	<% Else %>
	  <%=adoRS4.Fields("city_cd").Value %>
	<% End If %>
	</td>
    <td class="<%=strClass %>" height="20">Display</td>
    <td class="<%=strClass %>" height="20">
    <% If IsNull(adoRS4.Fields("shop_profile_desc").Value) Then %>
		All searched rates - no profile assoc.
	<% Else %>
	    <%=adoRS4.Fields("shop_profile_desc").Value %>
	<% End If %>
	</td>
   
  </tr>
  <%
        
        	adoRS4.MoveNext
        	
        Wend
        
   		adoRS4.Close
		Set adoRS4 = Nothing
		Set adoCmd4 = Nothing

  %>
    
    
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
  <p><font size="2">&nbsp;| <a href="javascript:maint_action(1);">Delete</a> 
  | <!-- <a href="javascript:maint_action(2)">Copy</a> | -->
  <a href="javascript:maint_action(3)">Enable</a> |
  <a href="javascript:maint_action(4)">Disable</font></a> |</font></p>
  <input type="hidden" name="refresh_from" value="search">
  <input type="hidden" name="action" value="1">
</form>
<form name="create_alert" method="POST" action="car_rate_rule_insert.asp" OnSubmit="CreateAlert();return false">
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
      <input type="text" name="rate_rule_id" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right; background-image:url('images/alt_color.gif')" value="<%=intRuleId %>" READONLY>&nbsp; <font face="Verdana" size="2">
      <input type="checkbox" name="copy" value="true">Save as a copy (leaves the 
      original unchanged)</font></td>
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
      <input type="text" name="alert_desc" size="20" style="width:439; font-family:Verdana; font-size:10pt; height:21" value="<%=strAlertDesc %>"></td>
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
                <input type="text" name="begin_dt" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')" value="<%=datBeginDate %>">
                <img src="images/cal.gif" id="dtimg1" style="position:relative" border="0" title="View Calendar" width="16" height="16" onclick="ShowCalendar(document.search_criteria.dtimg1, document.search_criteria.begin_dt, null, '<%=FormatDateTime(DateAdd("d", 1, Now),2) %>', '<%=FormatDateTime(DateAdd("d", 360, Now),2) %> - 1');event.cancelBubble=true;">
                </td>
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
                <input type="text" name="end_dt" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')" value="<%=datEndDate %>">
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
                <input type="text" name="first_pickup_dt" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')" value="<%=datFirstPickupDate %>">
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
      <font face="Verdana" size="2">(enter 'continuous' or blank for no pick-up date)</font></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">6. Last Pick-up:</font></td>
      <td width="672" height="25">
                <input type="text" name="last_pickup_dt" class="fsmall" style="width:200"  size='20' maxlength="10" onFocus="javascript:vDateType='1'" onKeyUp="DateFormat(this,this.value,event,false,'1')" onBlur="DateFormat(this,this.value,event,true,'1')" value="<%=datSecondPickupDate  %>">
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
      <font face="Verdana" size="2">(enter 'continuous' or blank for no pick-up date)</font></td>
      <td width="169" height="25">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="25">&nbsp;</td>
      <td width="247" height="25">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">7. System Rate 
      Code:</font></td>
      <td width="672" height="25">
      <input type="text" name="client_sys_rate_cd" size="20" style="width:200; font-family:Verdana; font-size:10pt" value="<%=strClientRateCode %>"></td>
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
      <select size="4" name="vend_cd1" multiple style="width:200; font-family:Verdana; font-size:10pt" >
      <% If strVendCd = "XX" Then %>
      <option selected value="XX"><%="All Competitors" %></option>
	  <% Else                     %>
      <option value="XX"><%="All Competitors" %></option>
      <% End  If                  %>

	 				<% While adoRS2.EOF = False %>
	 				<% If (InStr(1, strVendCd, adoRS2.Fields("vendor_cd").Value)) And (strVendCd <> "") Then %>
			              <option selected value="<%=adoRS2.Fields("vendor_cd").Value %>"><%=adoRS2.Fields("vendor_name").Value %></option>
	 				<% Else                     %>
			              <option value="<%=adoRS2.Fields("vendor_cd").Value %>"><%=adoRS2.Fields("vendor_name").Value %></option>
	 				<% End If                   %>
					<% 
						 adoRS2.MoveNext
					   Wend
					   Set adoRS2 = Nothing
					%>


      </select>


	              </select>      
      </td>
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
      <select size="4" name="vend_cd2" multiple style="width:200; font-family:Verdana; font-size:10pt; background-image:url('images/alt_color.gif'); background-repeat:repeat" >
      <% If strSelfCd = "XX" Then %>
      <option selected value="XX"><%="N/A" %></option>
	  <% Else                     %>
      <option value="XX"><%="N/A" %></option>
      <% End  If                  %>

	 				<% While adoRS3.EOF = False %>
	 				<% If (InStr(1, adoRS3.Fields("vendor_cd").Value, strSelfCd)) And (strVendCd <> "") Then %>
			              <option selected value="<%=adoRS3.Fields("vendor_cd").Value %>"><%=adoRS3.Fields("vendor_name").Value %></option>
	 				<% Else                     %>
			              <option value="<%=adoRS3.Fields("vendor_cd").Value %>"><%=adoRS3.Fields("vendor_name").Value %></option>
	 				<% End If                   %>
					<% 
						 adoRS3.MoveNext
					   Wend
					   Set adoRS3 = Nothing
					%>


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
      <select size="4" name="lor" style="width:200; font-family:Verdana; font-size:10pt">
      <% Select Case intLOR	%>
      
      <%	Case 1			%>
      <option selected value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      
      <%	Case 2			%>
      <option value="1">1</option>
      <option selected value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      
      <%	Case 3			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option selected value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>

      <%	Case 4			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option selected value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>

      <%	Case 5			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option selected value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      
      <%	Case 6			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option selected value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>

      <%	Case 7			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option selected value="7">7</option>
      <option value="8">8</option>

      <%	Case 8			%>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option selected value="8">8</option>

      
      <%	Case Else			%>
      <option selected value="1">1</option>
      <option value="5">5</option>
      
      
      <% End Select			%>
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
      <select size="4" name="city_cd" style="width:200; font-family:Verdana; font-size:10pt">
     
         <% While adoRS6.EOF = False                            %>
         <% 	If (InStr(strCityCd, adoRS6.Fields("city_cd").Value) = 0) Or (strCityCd = "") Then %>
         			<option value="<%=adoRS6.Fields("city_cd").Value %>"><%=adoRS6.Fields("city_cd").Value %></option>
         <% 	Else 											 %>		                    
         			<option selected value="<%=adoRS6.Fields("city_cd").Value %>"><%=adoRS6.Fields("city_cd").Value %></option>
						
		 <% 	End If 											 %>
		 <%    adoRS6.MoveNext 								     %>
		 <% Wend												 %> 
		 <% Set adoRS6 = Nothing 								 %>
      </select></td>
      <td width="169" height="71">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="24">&nbsp;</td>
      <td width="247" height="24">
      &nbsp;</td>
      <td width="200" height="24" valign="top" style="padding-top: 2px">&nbsp;</td>
      <td width="672" height="24">
      <font face="Verdana" size="2">(select 
      airport/city codes &quot;any&quot; for any location, edit to edit custom)</font></td>
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
      <select name="car_type_cd1" size="4" multiple style="width:200; font-family:Verdana; font-size:10pt">
	 <% If strCarTypeCd1 = "XXXX" Then %>
      <option selected value="XXXX"><%="N/A" %></option>
	  <% Else                     %>
      <option value="XXXX"><%="N/A" %></option>
      <% End  If                  %>


         <% While adoRS7.EOF = False                            %>
         <% 	If (InStr(strCarTypeCd1 , adoRS7.Fields("car_type_cd").Value) = 0) Or (strCarTypeCd1 = "") Then %>
         			<option value="<%=adoRS7.Fields("car_type_cd").Value %>"><%=adoRS7.Fields("car_type_cd").Value %></option>
         <% 	Else 											 %>		                    
         			<option selected value="<%=adoRS7.Fields("car_type_cd").Value %>"><%=adoRS7.Fields("car_type_cd").Value %></option>
						
		 <% 	End If 											 %>
		 <%    adoRS7.MoveNext 								     %>
		 <% Wend												 %> 
		 <% Set adoRS7 = Nothing 								 %>
      
      
<!--
      <option value="CCAR">CCAR</option>
      <option value="CFAR">CFAR</option>
      <option value="CPAR">CPAR</option>
      <option value="ECAR">ECAR</option>
      <option value="EDAR">EDAR</option>
      <option value="FCAR">FCAR</option>
      <option value="FDAR">FDAR</option>
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
-->
      </select> </font></td>
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
      <select name="car_type_cd2" size="4" multiple style="width:200; font-family:Verdana; font-size:10pt">
	 <% If strCarTypeCd2 = "XXXX" Then %>
      <option selected value="XXXX"><%="N/A" %></option>
	  <% Else                     %>
      <option value="XXXX"><%="N/A" %></option>
      <% End  If                  %>


         <% While adoRS8.EOF = False                            %>
         <% 	If (InStr(strCarTypeCd2 , adoRS8.Fields("car_type_cd").Value) = 0) Or (strCarTypeCd2 = "") Then %>
         			<option value="<%=adoRS8.Fields("car_type_cd").Value %>"><%=adoRS8.Fields("car_type_cd").Value %></option>
         <% 	Else 											 %>		                    
         			<option selected value="<%=adoRS8.Fields("car_type_cd").Value %>"><%=adoRS8.Fields("car_type_cd").Value %></option>
						
		 <% 	End If 											 %>
		 <%    adoRS8.MoveNext 								     %>
		 <% Wend												 %> 
		 <% Set adoRS8 = Nothing 								 %>
      </select> </font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19">&nbsp;</td>
      <td width="672" height="19">
      <font face="Verdana" size="2">(use &quot;any&quot; to 
      denote any car type, you use &quot;any&quot; for<br>
&nbsp;both items to have the system to match car types)</font></td>
      <td width="169" height="19">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19"><font face="Verdana" size="2">12. Rate Source::</td>
      <td width="672" height="19">
      <select size="1" name="data_source" style="width:200; font-family:Verdana; font-size:10pt">
	  <option selected value="ALL">Determined by profile</option>      
<!--

      <% Select Case strDataSource	%>
      
      <%	Case "CAR"	%>
	      <option selected value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "ORB"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option selected value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "EXP"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option selected value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "TRV"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option selected value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>

      <%	Case "ATV"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option selected value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>
      
      <%	Case "ALL"	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option selected value="ALL">All sites</option>


      <%	Case Else	%>
	      <option value="CAR">Brand sites (avis.com, etc.)</option>
	      <option selected value="ORB">Orbitz</option>
	      <option value="EXP">Expedia</option>
	      <option value="TRV">Travelocity</option>
	      <option value="ATV">All Travel sites</option>
	      <option value="ALL">All sites</option>


      <% End Select 				%>
-->
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
      <select size="1" name="situation_cd" style="width:200; font-family:Verdana; font-size:10pt; height:24">
      
      <% If intSituationCd = 0 Then	%>
	      <option selected value="0">(None selected)</option>
	  <% Else %>
	      <option value="0">(None selected)</option>
	  <% End If %>
	  
      <% If intSituationCd = 1 Then	%>
	      <option selected value="1">> (all competitors)</option>
	  <% Else %>
	      <option value="1">> (all competitors)</option>
	  <% End If %>
	  
      <% If intSituationCd = 2 Then	%>
	      <option selected value="2">> (any competitor)</option>
	  <% Else %>
	      <option value="2">> (any competitor)</option>
	  <% End If %>
	  
      <% If intSituationCd = 3 Then	%>
	      <option selected value="3">> (custom)</option>
	  <% Else %>
	      <option value="3">> (custom)</option>
	  <% End If %>
	  
      <% If intSituationCd = 4 Then	%>
	      <option selected value="4">&gt;= (all competitors)</option>
	  <% Else %>
	      <option value="4">&gt;= (all competitors)</option>
	  <% End If %>
	  
      <% If intSituationCd = 5 Then	%>
	      <option selected value="5">>= (any competitor)</option>
	  <% Else %>
	      <option value="5">>= (any competitor)</option>
	  <% End If %>
	  
      <% If intSituationCd = 6 Then	%>
	      <option selected value="6">>= (custom)</option>
	  <% Else %>
	      <option value="6">>= (custom)</option>
	  <% End If %>
	  
      <% If intSituationCd = 7 Then	%>
	      <option selected value="7">= (all competitors)</option>
	  <% Else %>
	      <option value="7">= (all competitors)</option>
	  <% End If %>

      <% If intSituationCd = 8 Then	%>
	      <option selected value="8">= (any competitor)</option>
	  <% Else %>
	      <option value="8">= (any competitor)</option>
	  <% End If %>

      <% If intSituationCd = 9 Then	%>
	      <option selected value="9">= (custom)</option>
	  <% Else %>
	      <option value="9">= (custom)</option>
	  <% End If %>

      <% If intSituationCd = 11 Then	%>
	      <option selected value="11">< = (all competitors)</option>
	  <% Else %>
	      <option value="11">< = (all competitors)</option>
	  <% End If %>

      <% If intSituationCd = 12 Then	%>
	      <option selected value="12"><= (any competitor)</option>
	  <% Else %>
	      <option value="12"><= (any competitor)</option>
	  <% End If %>
	  
      <% If intSituationCd = 13 Then	%>
	      <option selected value="13">&lt;= (custom)</option>
	  <% Else %>
	      <option value="13">&lt;= (custom)</option>
	  <% End If %>
	  
      <% If intSituationCd = 14 Then	%>
	      <option selected value="14">Less than (all) by amt</option>
	  <% Else %>
	      <option value="14">Less than (all) by amt</option>
	  <% End If %>
	  
      <% If intSituationCd = 16 Then	%>
	      <option selected value="16">&lt; (any competitor)</option>
	  <% Else %>
	      <option value="16">&lt; (any competitor)</option>
	  <% End If %>

      <% If intSituationCd = 15 Then	%>
	      <option selected value="15">< (custom)</option>
	  <% Else %>
	      <option value="15">< (custom)</option>
	  <% End If %>
	  
      <% If intSituationCd = 17 Then	%>
	      <option selected value="17">&lt;&gt; (any competitor) - diff</option>
	  <% Else %>
	      <option value="17">&lt;&gt; (any competitor) - diff</option>
	  <% End If %>
	  
      <% If intSituationCd = 18 Then	%>
	      <option selected value="18">&lt;&gt; (all competitors) - diff</option>
	  <% Else %>
	      <option value="18">&lt;&gt; (all competitors) - diff</option>
	  <% End If %>
	  
      <% If intSituationCd = 19 Then	%>
	      <option selected value="19">&lt;&gt; (any competitor) + diff</option>
	  <% Else %>
	      <option value="19">&lt;&gt; (any competitor) + diff</option>
	  <% End If %>
	  
      <% If intSituationCd = 19 Then	%>
	      <option selected value="20">&lt;&gt; (all competitors) + diff</option>
	  <% Else %>
	      <option value="20">&lt;&gt; (all competitors) + diff</option>
	  <% End If %>
 


      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="23">&nbsp;</td>
      <td width="247" height="23">&nbsp;</td>
      <td width="200" height="23"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      by amount:<br>
&nbsp;&nbsp;&nbsp; <br>
&nbsp;</font></td>
      <td width="672" height="23">
      <input type="text" name="situation_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strSituationAmt %>">
      <br>
      <input type="radio" value="1" name="is_dollar" style="font-family:Verdana; font-size:10pt" checked><font face="Verdana" size="2">Whole Dollar<br>
      <input type="radio" value="0" name="is_dollar" style="font-family:Verdana; font-size:10pt"><font face="Verdana" size="2">Percentage</td>
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
      <td width="200" height="22"><font face="Verdana" size="2">14. Quantity &amp; 
      Period:</font></td>
      <td width="672" height="22">
      <select size="1" name="quantity_period_cd" style="width:200; font-family:Verdana; font-size:10pt">
      <% If intQuantityPeriodCd = 0 Then %>
	      <option selected value="0">Each time it occurs</option>
	  <% Else %>
	      <option value="0">Each time it occurs</option>
	  <% End If %>
      
      <% If intQuantityPeriodCd = 1 Then %>
	      <option selected  value="1">After X Events</option>
	  <% Else %>
	      <option value="1">After X Events</option>
	  <% End If %>
	  
      <% If intQuantityPeriodCd = 2 Then %>
	      <option selected value="2">After X Events in X Hours</option>
	  <% Else %>
	      <option value="2">After X Events in X Hours</option>
	  <% End If %>

      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      No. of Events:</font></td>
      <td width="672" height="22">
      <input type="text" name="event_count" size="20" style="width:200; font-family:Verdana; font-size:10pt; text-align:right" value="<%=intEventCount %>"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      That happen in:</font></td>
      <td width="672" height="22">
      <input type="text" name="event_time_count" size="20" style="width:137; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=intEventTimeCount %>">
      <select size="1" name="event_time_type">
      <% If strEventTimeType  = "Hours" Or strEventTimeType = "" Then %>
      <option selected>Hours</option>
      <option>Days</option>
	  <% Else %>
      <option>Hours</option>
      <option selected >Days</option>
	  <% End If %>

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
      <td width="200" height="22"><font face="Verdana" size="2">15. Response:</font></td>
      <td width="672" height="22">
      <select size="1" name="response_cd" style="width:200; font-family:Verdana; font-size:10pt" onclick="CheckResponse()" >
	  <% If intResponseCd = 0 Then	%>
      <option selected value="0">(None selected)</option>
	  <% Else						%>
      <option value="0">(None selected)</option>
	  <% End If 					%>	  
      
	  <% If intResponseCd = 1 Then	%>
      <option selected value="1">Competitor&#39;s Rate minus</option>
	  <% Else						%>
      <option value="1">Competitor&#39;s Rate minus</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 2 Then	%>
      <option selected value="2">Competitor&#39;s Rate plus</option>
	  <% Else						%>
      <option value="2">Competitor&#39;s Rate plus</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 3 Then	%>
      <option selected value="3">Raise Rate - to amount</option>
	  <% Else						%>
      <option value="3">Raise Rate - to amount</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 4 Then	%>
      <option selected value="4">Raise Rate - by percentage</option>
	  <% Else						%>
      <option value="4">Raise Rate - by percentage</option>
	  <% End If 					%>	  

	  <% If intResponseCd = 5 Then	%>
      <option selected value="5">Raise Rate - by amount</option>
	  <% Else						%>
      <option value="5">Raise Rate - by amount</option>
	  <% End If 					%>	  

      </select>
      </td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="19">&nbsp;</td>
      <td width="247" height="19">&nbsp;</td>
      <td width="200" height="19"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Amount:<br>
      <br>&nbsp;</font></td>
      <td width="672" height="19">
      <input type="text" name="response_amt" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strResponseAmt %>" >
      <br>
      <% If blnResponseDollar Then %>
	      <input type="radio" value="1" name="is_response_dollar" style="font-family:Verdana; font-size:10pt" checked><font face="Verdana" size="2">Whole Dollar<br>
	      <input type="radio" value="0" name="is_response_dollar" style="font-family:Verdana; font-size:10pt">Percentage</td>
	  <% Else					   %>
	      <input type="radio" value="1" name="is_response_dollar" style="font-family:Verdana; font-size:10pt" ><font face="Verdana" size="2">Whole Dollar<br>
	      <input type="radio" value="0" name="is_response_dollar" style="font-family:Verdana; font-size:10pt" checked>Percentage</td>
	  <% End If					   %>
	  
	  
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
      <td width="200" height="22"><font face="Verdana" size="2">16. Search Type:</font></td>
      <td width="672" height="22">
      <select size="1" name="search_profile" style="width:200; font-family:Verdana; font-size:10pt">
      <option selected value="1">Link to Profile (Required)</option>

<!--      
      <% If blnSearchProfile = 0 Then	%>
      <option selected value="0">As searched (all searches)</option>
      <option value="1">Link to Profile</option>
      <% Else						%>
      <option value="0">As searched (all searches)</option>
      <option selected value="1">Link to Profile</option>
      <% End If						%>
-->      
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
      Profile(s):<br>
      <br>&nbsp;</font></td>
      <td width="672" height="22">
		<select name="profile_id" size="4" multiple style="width:373; font-family:Verdana; font-size:10pt; height:70">
                <% While adoRS9.EOF = False %> 
                	<% If adoRS9.Fields("profile_id").Value = 0 Then %>
	                	<% adoRS9.MoveNext %>
	                <% end If %>
                	<% If InStr(1, strProfileID, adoRS9.Fields("profile_id").Value) > 0 Then %>
	                	<option selected value="<%=adoRS9.Fields("profile_id").Value %>"><%=adoRS9.Fields("desc").Value %></option>
	                <% Else %>
		               <% If adoRS9.Fields("profile_status").Value = "E" Then %>
		                	<option value="<%=adoRS9.Fields("profile_id").Value %>"><%=adoRS9.Fields("desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS9.MoveNext
				   Wend
					   
				   Set adoRS9 = Nothing
				%>
				</select>	</td>
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
      <font face="Verdana" size="2">17. Rate Range:</font></td>
      <td width="672" height="22"><font face="Verdana" size="2">
      (leave blank or enter a zero to indicate no limit)</font></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
      Maximum:</font></td>
      <td width="672" height="22">
      <input type="text" name="rate_maximum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:21" value="<%=strRangeMax %>"></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">
      &nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; 
      Minimum:</font></td>
      <td width="672" height="25">
      <input type="text" name="rate_minimum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strRangeMin %>"></td>
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
      <td width="200" height="22"><font face="Verdana" size="2">18. Evaluate on 
      Success:</font></td>
      <td width="672" height="22">
      <select size="1" name="on_success_id" style="width:370; font-family:Verdana; font-size:10pt; height:24">
      <option selected value="0">(No follow-on rule)</option>
      
                <% While adoRS18a.EOF = False %> 
                	<% If intSuccessId = adoRS18a.Fields("rate_rule_id").Value Then %>
	                	<option selected value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
	                <% Else %>
		               <% If adoRS18a.Fields("rule_status").Value = "E" Then %>
		                	<option value="<%=adoRS18a.Fields("rate_rule_id").Value %>"><%=adoRS18a.Fields("alert_desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS18a.MoveNext
				   Wend
					   
				   Set adoRS18a = Nothing
				%>
      
      </select></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      Evaluate on Failure:</font></td>
      <td width="672" height="22">
		<select name="on_failure_id" style="width:370; font-family:Verdana; font-size:10pt; height:24" size="1">
     <option selected value="0">(No follow-on rule)</option>
      
                <% While adoRS18b.EOF = False %> 
                	<% If intFailureId = adoRS18b.Fields("rate_rule_id").Value Then %>
	                	<option selected value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
	                <% Else %>
		               <% If adoRS18b.Fields("rule_status").Value = "E" Then %>
		                	<option value="<%=adoRS18b.Fields("rate_rule_id").Value %>"><%=adoRS18b.Fields("alert_desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS18b.MoveNext
				   Wend
					   
				   Set adoRS18a = Nothing
				%>
				</select>	</td>
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
      <td width="200" height="25"><font face="Verdana" size="2">
      19. Utilization Range:</td>
      <td width="672" height="25"><font face="Verdana" size="2">
      (leave blank or enter a zero to indicate no limit)</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22"><font face="Verdana" size="2">
      &nbsp;&nbsp;&nbsp;&nbsp; Maximum:</td>
      <td width="672" height="22">
      <font face="Verdana" size="2">
      <input type="text"  name="utilization_maximum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right" value="<%=strUtilMax %>"></font></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="25"><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp;&nbsp; Minimum:</td>
      <td width="672" height="25">
      <font face="Verdana" size="2">
      <input type="text" name="utilization_minimum" size="20" style="width:127; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=strUtilMin %>"></font></td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">
      &nbsp;</td>
</font>
      <td width="672" height="22">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
          <input name="<%=strButton %>" type="submit" id="submit" value="    <%=strButton %>   " class="rh_button"></font></td>
      <td width="672" height="22">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="22">&nbsp;</td>
      <td width="247" height="22">&nbsp;</td>
      <td width="200" height="23">&nbsp;</td>
      <td width="672" height="23">
      &nbsp;</td>
      <td width="169" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="11" height="61">&nbsp;</td>
      <td width="247" height="61">&nbsp;</td>
      <td width="872" colspan="2" height="61">
      &nbsp;</td>
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
  <input type="hidden" name="rule_status" value="E">
</form>
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
<%	
	Rem Clean up
	Set adoCmd = Nothing 
	Set adoCmd1 = Nothing 
	Set adoCmd2 = Nothing 
	Set adoCmd3 = Nothing 
	
%>

<script language="javascript">
	document.search_alerts.name.focus();
</script>