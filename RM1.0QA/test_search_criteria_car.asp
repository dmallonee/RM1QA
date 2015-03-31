<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180


	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd1	
	Dim adoRS1
	Dim adoCmd2	
	Dim adoRS2

	Dim adoPrices
	Dim strUserId

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	strSelfVendCd = Request.Cookies("rate-monitor.com")("vend_cd")
		
	strConn = Session("pro_con")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_profile_select"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Null)
		
	Set adoRS = adoCmd.Execute
		
	Rem Get the data sources
	Set adoCmd1 = CreateObject("ADODB.Command")

	adoCmd1.ActiveConnection =  strConn
	adoCmd1.CommandText = "data_source_select"
	adoCmd1.CommandType = 4

	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@lob_id", 3, 1, 0, 2)
		
	Set adoRS1 = adoCmd1.Execute

	Rem Get the vendors
	Set adoCmd2 = CreateObject("ADODB.Command")

	adoCmd2.ActiveConnection =  strConn
	adoCmd2.CommandText = "vendor_select"
	adoCmd2.CommandType = 4
		
	Set adoRS2 = adoCmd2.Execute
	
	Rem Get the vendors
	Set adoCmd3 = CreateObject("ADODB.Command")

	adoCmd3.ActiveConnection =  strConn
	adoCmd3.CommandText = "user_city_select"
	adoCmd3.CommandType = 4
	
	adoCmd3.Parameters.Append adoCmd3.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS3 = adoCmd3.Execute
	
	Rem Get the car types
	Set adoCmd5 = CreateObject("ADODB.Command")

	adoCmd5.ActiveConnection =  strConn
	adoCmd5.CommandText = "car_type_select"
	adoCmd5.CommandType = 4
		
	Set adoRS5 = adoCmd5.Execute
		
	
	If (Request("profile") > 0) Then
	


		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_select"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Request("profile"))

		Set adoRS4 = adoCmd.Execute
		ProfileID = adoRS4.Fields("profile_id").Value
		
		blnProfileLoad = True

	Else

		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_select"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, 0)

		Set adoRS4 = adoCmd.Execute
		ProfileID = adoRS4.Fields("profile_id").Value
	
		blnProfileLoad = False
		
	End If


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Criteria</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<!-- European format dd-mm-yyyy -->
<!--
<script language="JavaScript" src="inc/calendar1.js"></script>
-->
<!-- Date only with year scrolling -->
<!-- American format mm/dd/yyyy -->
<!--
<script language="JavaScript" src="inc/calendar2.js"></script>
-->
<!-- Date only with year scrolling -->
<script language="JavaScript" src="inc/ts_picker.js"></script>
<script language="JavaScript" src="inc/center.js"></script>
<script language="JavaScript" src="inc/date_check.js"></script>
<script language="JavaScript" src="inc/multiple_select_support.js"></script>
<script language="JavaScript" src="inc/multiple_select_support2.js"></script>
<script language="JavaScript" src="inc/selectbox.js"></script>
<script language="Javascript"> 

function advanced_search() {
	alert("Advanced Search is currently off-line" + "\n" + "Reason: Unable to detect Sabre connection");
	//return true;
}


function not_enabled() {
	alert("This section is currently  unavailable." + "\n" + "Please contact your Rate-Highway rep if you would like to enable it");
	//return true;
}


function update_company_list() { 

	var list = new String("")
	
	removeAllOptions(document.search_criteria.highlighted_company);
	addOption(document.search_criteria.highlighted_company, "(please select one)", "", true);
	
	for (i = 0; i < document.search_criteria.selected_companies.options.length;i++)
	{
		addOption(document.search_criteria.highlighted_company, document.search_criteria.selected_companies.options[i].text, document.search_criteria.selected_companies.options[i].value, false);
		if (i == 0)
		{
	      list = document.search_criteria.selected_companies.options[i].value;
		}
		else
		{
	      list = list + ',' + document.search_criteria.selected_companies.options[i].value;
		}
	}

    document.search_criteria.company_list.value = list; 
    //document.search_criteria.selected_companies.options.text; 
 
} 
function update_car_type_list() { 

	var list = new String("")
	
	for (i = 0; i < document.search_criteria.selected_car_types.options.length;i++)
	{
		if (i == 0)
		{
	      list = document.search_criteria.selected_car_types.options[i].value
		}
		else
		{
	      list = list + ',' + document.search_criteria.selected_car_types.options[i].value
		}
	}

    document.search_criteria.car_type_list.value = list 
    //document.search_criteria.selected_companies.options.text; 
     
} 

function update_city_list() { 

	var list = new String("")
	
	for (i = 0; i < document.search_criteria.selected_cities.options.length;i++)
	{
		if (i == 0)
		{
	      list = document.search_criteria.selected_cities.options[i].value
		}
		else
		{
	      list = list + ',' + document.search_criteria.selected_cities.options[i].value
		}
	}


    document.search_criteria.cities_list.value = list 
    return true 
} 


function check_all_days(fieldName) {

    thisButton = document.forms[0][fieldName]; 
    for( var i=0; i<thisButton.length; i++ ) { 
        thisButton[i].checked = true; 
    } 

}

</script>
<!--
<iframe src="calb.htm" style="display:none;position:absolute;width:148;height:194;z-index=100" id="CalFrame" marginheight="0" marginwidth="0" noresize frameborder="0" scrolling="NO">
</iframe>
<script language="JavaScript">

//
// Expedia Style Calendar Control Scripts
//



var cF=document.all.CalFrame;var cW=window.frames.CalFrame;var g_tid=0;var g_cP,g_eD,g_eDP,g_dmin,g_dmax,g_htm;

function CB(){event.cancelBubble=true}

function SCal(cP,eD,eDP,dmin,dmax,htm){
	clearTimeout(g_tid);
	var s=(g_eD==eD);
	g_cP=cP;
	g_eD=eD;
	g_eDP=eDP;
	g_dmin=dmin;
	g_dmax=dmax;
	g_htm=htm;
	WaitCal(true,s);
	}
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
	var htm="save_cal.htm";
	SCal(eP,eD,eDP,dmin,dmax,htm);
}

</script>

<script for="document" event="onclick()">
-->
<!--
CancelCal();
//-->
</script>
<script language="JavaScript" type="text/JavaScript">

function CopyCityCode()
{
	//document.search_criteria.return_city.value = document.search_criteria.pickup_city.options[document.search_criteria.pickup_city.selectedindex].value;
	selected   = document.search_criteria.pickup_city.selectedIndex; 
    fieldValue = document.search_criteria.pickup_city.options[selected].value; 
	document.search_criteria.return_city.value = fieldValue ;
	
	return;
}
</script>

<script type='text/javascript' language='javascript' >
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//
// Page submition section
//
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function confirmSubmit(SubmitType) {

	var selected
	var fieldValue
	var fieldText

	if (SubmitType == 'open'){
		
		//verify that a profile has been selected
	   	selected   = document.search_criteria.profile.selectedIndex;
	   	fieldValue = document.search_criteria.profile.options[selected].value;
	   	fieldText  = document.search_criteria.profile.options[selected].text;

   		if( fieldValue == 0 ) { 
	        alert( "You must select a valid profile before you can open it"); 
	        return false;

	    	} 
		else {
			document.search_criteria.action = 'search_criteria_car.asp';
			//document.search_criteria.submit;
			return true;   
			} 
		}

	else if (SubmitType == 'save'){
		
		// This needs to be cleared out so that it is not confused with a save as
		document.search_criteria.profile_save_as.value = "" 

		//verify that a profile has been selected
	   	selected   = document.search_criteria.profile.selectedIndex; 
	   	fieldValue = document.search_criteria.profile.options[selected].value;
	   	fieldText  = document.search_criteria.profile.options[selected].text;
 
   		if( fieldValue == 0 ) { 
	        alert( "You must select and open a profile before you can save it"); 
	        return false;

	    	} 

	
		else {
			if (confirm("Are you sure you want to overwrite " + fieldText  + "?")) {
				if (ValidateForm() == true) {
					//alert("ValidateForm() succeded");
					document.search_criteria.action = 'search_profile_insert_car.asp';
					//document.search_criteria.submit();					
					return true;
					}
				else {
					//alert("ValidateForm() failed");
					return false;
					}   
				}  
			else {  
				//alert("passed on overwriting");
				return false;   
				} 
			}
		}

	else if (SubmitType == 'saveas'){
		
		//verify that a profile has been named
	    validChars  = "abcdefghijklmnopqrstuvwxyz"; 
	    validChars += "ABCDEFGHIJKLMNOPQRSTUVWXYZ"; 
	    validChars += "0123456789"; 
 
	    fieldName   = document.search_criteria.profile_save_as; 
	    fieldValue  = fieldName.value; 
	    fieldLength = fieldValue.length; 
	    minLength   = 1; 
	    maxLength   = 255; 
 
	    var err01   = "Non-valid character(s) found in the profile save as name, please no special characters."; 
	    var err03   = "Please enter a profile name with at least " + minLength + " character in length."; 
	    var err04   = "Please enter less than " + maxLength + " characters in length for the profile name."; 

	    if( fieldValue == "" ) { 
	        alert( err03 ); 
	        fieldName.focus(); 
	        return false;
	    	} 
	    else if ( fieldLength < minLength ) { 
	        alert( err03 ); 
	        fieldName.focus();
	        return false; 
	    	} 
	    else if (( fieldLength > maxLength ) && ( maxLength > 0 )) { 
	        alert( err04 ); 
	        fieldName.focus();
	        return false; 
	    	} 
	    	
	    else { 
	        for( var i=0; i<fieldLength; i++ ) { 
	            if ( validChars.indexOf( fieldValue.charAt( i )) == -1 ) { 
	                alert( err01 ); 
	                fieldName.focus(); 
	                return false; 
	            	} 
	            else { 
					if (ValidateForm() == true) {
						document.search_criteria.action='search_profile_insert_car.asp';
						//document.search_criteria.submit();
						return true;   
		            	} 
					else {  
						return false;   
						} 

		        	} 
		    	} 
			} 
		}

	else if (SubmitType == 'searchnow'){
		
		if (confirm("Are you sure you want to search?")) {  
			if (ValidateForm() == true) {
				document.search_criteria.action='search_request_insert_car.asp';
				//document.search_criteria.submit();
				return true;   
				}  
			else {  
				return false;   
				} 
			}
		}

}

//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//
// Page validation section
//
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

function ValidateForm() {

    doSubmit=new Boolean()
    doSubmit=true
    

	if (doSubmit==true) {
	    if (document.search_criteria.selected_cities.options.length < 1 ) {
	        alert("Please select at least one city code to search.");
	        document.search_criteria.selected_cities.focus();
	        doSubmit=false;
	    }
	}


	if (doSubmit==true) {
	    //alert("entered ValidateForm");
	    //DOW list
	    selection  = null; 
	    thisButton = document.search_criteria.dow_list; 
	    for( var i=0; i<thisButton.length; i++ ) { 
	        if( thisButton[i].checked ) { 
	           selection = thisButton[i].value; 
	        } 
	    } 

	    if( selection == null ) { 
	        alert( "Please check at least one of the day of week boxes." ); 
	        doSubmit=false

	    } 
	}

   //selected   = document.forms[0][fieldName].selectedIndex; 
   // fieldValue = document.forms[0][fieldName].options[selected].value; 
 
   // if( fieldValue == "" ) { 
   //     alert( "'" + fieldLabel +  "' - Select one of the options." ); 
   // } 

    /*
	if (doSubmit==true) {
	    if (document.search_criteria.return_city.value=="") {
	        alert("Please select a valid city code");
	        document.search_criteria.pickup_city.focus();
	        doSubmit=false;
	    }
	}
    */
  
	if (doSubmit==true) {
	    if (document.search_criteria.begin_date.value=="") {
	        alert("Please enter a valid begin date");
	        document.search_criteria.begin_date.focus();
	        doSubmit=false;
	    }
	}


	if (doSubmit==true) {
	    if (document.search_criteria.end_date.value=="") {
	        alert("Please enter a valid end date");
	        document.search_criteria.end_date.focus();
	        doSubmit=false;
	    }
	}

	if (doSubmit==true) {
	    if (document.search_criteria.selected_car_types.options.length==0) {
	        alert("Please select at least one car type");
	        document.search_criteria.unselected_car_types.focus();
	        doSubmit=false;
	    }
	}

	if (doSubmit==true) {
	    if (document.search_criteria.selected_companies.options.length==0) {
	        alert("Please select at least one car company");
	        document.search_criteria.unselected_companies.focus();
	        doSubmit=false;
	    }
	}

	if (doSubmit==true) {
	    selected   = document.search_criteria.highlighted_company.selectedIndex; 
	    fieldValue = document.search_criteria.highlighted_company.options[selected].value; 
 
	    if( fieldValue == "" ) { 
	        alert( "Please select a company for highlight" ); 
	        document.search_criteria.highlighted_company.focus();
	        doSubmit=false;
	    }
	} 


    //alert("leaving ValidateForm");
    return doSubmit
    
 
}

function verifyDateOrder() {
			var begin_date = new Date(document.search_criteria.begin_date)
			var end_date = new Date(document.search_criteria.end_date)
			if (begin_date < end_date) {
	        	alert("The 1st pick-up date must be the same or prior to the 2nd pick-up date");
	        	document.search_criteria.end_date.focus();
	        	}
	        }
	        



function two() {
var o_parent;
var o_cal;
var x_box;
var y_box;

	o_parent = document.getElementById("rental_begin_date");
	
report.value = "Each object has a set of offset positions.  The offset properties for the positioned darkBlue DIV above are: \n";

	while (o_parent.offsetParent.tagName != "BODY") {
		report.value= report.value + "   offsetLeft = " + o_parent.offsetLeft + "\n";
		report.value= report.value + "   offsetTop = " + o_parent.offsetTop + "\n";
		report.value= report.value + "   offsetHeight = " + o_parent.offsetHeight + "\n";
		report.value= report.value + "   offsetWidth = " + o_parent.offsetWidth + "\n";
		report.value= report.value + "   tagName = " + o_parent.tagName + "\n";
		o_parent = o_parent.offsetParent;
	}



//report.value= report.value + "   fromrowed.offsetLeft = " + fromrowed.offsetLeft + "\n";
//report.value= report.value + "   fromrowed.offsetTop = " + fromrowed.offsetTop + "\n";
//report.value= report.value + "   fromrowed.offsetHeight = " + fromrowed.offsetHeight + "\n";
//report.value= report.value + "   fromrowed.offsetWidth = " + fromrowed.offsetWidth + "\n";

report.value= report.value + "   document.body.offsetLeft = " + document.body.offsetLeft + "\n";
report.value= report.value + "   document.body.offsetTop = " + document.body.offsetTop + "\n";
report.value= report.value + "   document.body.offsetHeight = " + document.body.offsetHeight + "\n";
report.value= report.value + "   document.body.offsetWidth = " + document.body.offsetWidth + "\n";

	x_box = document.getElementById("x");
	y_box = document.getElementById("y");

	o_cal = document.getElementById("test_div");
	o_cal.style.top =  y_box.value + "px" : "0px";
	o_cal.style.left = x_box.value + "px";


}

</script>





<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all">
<script language="javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>	
<style>
<!--
.testdiv     { position: absolute; left: 200; top: 200; width: 100; height: 100; 
               background-color: #00FF00 }
-->
</style>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">

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
                <td>
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
        <td><img src="images/h_search_criteria.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<!--
<form method="POST" name="search_criteria" id="1" action="" onSubmit="return submitForm()" >
-->
<form method="POST" name="search_criteria" id="1" action="" >
  <!-- JUSTTABS TOP OPEN --><br>
  <table cellpadding="0" cellspacing="0" border="0" align="CENTER" width="770" bgcolor="#FFFFFF">
    <tr height="1">
      <td colspan="1" width="10">&nbsp;</td>
      <td rowspan="2" width="280"><a href="search_criteria_car.asp">
      <img src="images/quicksearchwebonly0_a.GIF" width="161" height="25" hspace="0" vspace="0" border="0" alt="QuickSearch (web only)" description="Quick Search"></a><a href="archive/search_criteria_car_adv.asp"><img src="images/advancedsearch1_ia.GIF" width="119" height="25" hspace="0" vspace="0" border="0" alt="Advanced Search" description="Advanced Search" ></a></td>
      <td colspan="1">&nbsp;</td>
    </tr>
    <tr height="1">
      <td bgcolor="#000000" colspan="1" height="1">
      <img src="pixel.gif" width="1" height="1"></td>
      <td bgcolor="#000000" colspan="1" height="1">
      <img src="pixel.gif" width="1" height="1"></td>
    </tr>
  </table>
   <table align="center" cellpadding="0" cellspacing="0" border="0" width="769" bgcolor="#CCCCFF" name="criteria" id="criteria">
      <tr>
        <td width="1" bgcolor="#000000">
        <img src="pixel.gif" width="1" height="1"></td>
        <td bgcolor="#D9DEE1" width="768">
        <table border="0" cellspacing="5" cellpadding="5" width="768" name="inner_criteria">
          <tr>
            <td bgcolor="#FFFFFF" width="748"><font color="#080000">
            <!-- JUSTTABS TOP CLOSE -->
            <table width="745" border="0" cellspacing="0" cellpadding="2">
              <tr valign="bottom">
                <td width="162" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="93" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="5" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="52" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="83" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="159" bgcolor="#FFFFFF">&nbsp;</td>
                <td bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td width="162" bgcolor="#FFFFFF">
                <img src="images/ti_profile.gif" width="162" height="25"></td>
                <td width="93" bgcolor="#FFFFFF">
                <button name="open" style="height: 25px; width: 90px" value="open" onclick="return confirmSubmit('open')" type="submit">
                Open</button></td>
                <td width="5" bgcolor="#FFFFFF">
                <button name="save" style="height: 25; width: 90" value="save" onclick="return confirmSubmit('save')" type="submit">
                Save</button></td>
                <td width="52" bgcolor="#FFFFFF"></td>
                <td width="83" bgcolor="#FFFFFF"></td>
                <td width="159" bgcolor="#FFFFFF">
                <button name="save_as" style="height: 25; width: 90" value="save_as" type="submit" onclick="return confirmSubmit('saveas')">
                Save As</button></td>
                <td bgcolor="#FFFFFF">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td width="162" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="150" colspan="3" bgcolor="#FFFFFF">
                <select name="profile" style="width:230;">
                <option value="0">Default</option>
                <% While adoRS.EOF = False %> 
	                <% If ProfileID = adoRS.Fields("profile_id").Value Then %>
	                	<option selected value="<%=adoRS.Fields("profile_id").Value %>"><%=adoRS.Fields("desc").Value %></option>
	                <% Else %>
		               <% If adoRS.Fields("profile_status").Value = "E" Then %>
		                	<option value="<%=adoRS.Fields("profile_id").Value %>"><%=adoRS.Fields("desc").Value %></option>
		                <% End If %> 
					<% End If %>                
                <%	adoRS.MoveNext
				   Wend
					   
				   Set adoRS = Nothing
				%>
				</select>
				</td>
                <td width=" 83" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="316" colspan="2" bgcolor="#FFFFFF">
                <input type="text" name="profile_save_as" size="36" onfocus="this.className='focus';cl(this,'save profile as...');" onblur="this.className='';fl(this,'save profile as...');"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td bgcolor="#FFFFFF">
                <img src="images/separator.gif" width="20" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="2">
              <tr valign="bottom">
                <td width="163" bgcolor="#FFFFFF" valign="top">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <img src="images/ti_city_location.gif" width="162" height="25"></font></td>
                <td width="20" valign="top">
                <img src="images/separator.gif" width="20" height="20"></td>
                <td width="543" valign="bottom">
                <table width="475" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table1">
                  <tr>
                    <td width="475" colspan="3">
                      <div align="center">
                        <font size="2">(use buttons to move between selected/unselected)
                        Unselected cities&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        Selected cities&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; </font>
                      </div>
                    
                    </td>
                  </tr>
                  <tr>
                    <td width="209">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="unselected_cities" size="5" style="width:200;" multiple >
                      <% Dim Cities                                          %> 
                      <% Dim SelectedCities  							     %>
                      <% Cities = adoRS4.Fields("city_cd").Value     		 %> 
                      <% While adoRS3.EOF = False                            %>
                      <% 	If (InStr(Cities, adoRS3.Fields("city_cd").Value) = 0) Or (Cities = "") Then %>
                      			<option value="<%=adoRS3.Fields("city_cd").Value %>"><%=adoRS3.Fields("city_cd").Value %></option>
                      <% 	Else 											 %>		                    
		              <%     	SelectedCities = SelectedCities & TRIM(adoRS3.Fields("city_cd").Value) & ","  %>
						
					  <% 	End If 											 %>
					  <%    adoRS3.MoveNext 								 %>
					  <% Wend												 %> 
					  <% Set adoRS3 = Nothing 								 %>
					</select></font></td>
                    <td width="37">
                      <img border="0" src="images/move_right.GIF" width="23" height="22"  onclick="moveDualList( document.search_criteria.unselected_cities, document.search_criteria.selected_cities, false );update_city_list();return false" ><br>
                      <img border="0" src="images/move_right_all.GIF" width="23" height="22"  onclick="moveDualList( document.search_criteria.unselected_cities, document.search_criteria.selected_cities, true );update_city_list();return false"  ><br>
                      <img border="0" src="images/move_left.GIF" width="23" height="22"  onclick="moveDualList( document.search_criteria.selected_cities, document.search_criteria.unselected_cities, false );update_city_list();return false"  ><br>
                      <img border="0" src="images/move_left_all.GIF" width="23" height="22"  onclick="moveDualList( document.search_criteria.selected_cities, document.search_criteria.unselected_cities, true );update_city_list();return false"  >
                    </td>
                    <td width="228">
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                      <select name="selected_cities" size="5" style="width:200;" multiple>
                      <% Dim CityArray                                            %> 
                      <% CityArray = Split(SelectedCities, ",")                   %>
                      <% For intIndex = 0 To UBound(CityArray) - 1                %>
                        <option value="<%=CityArray(intIndex) %>"><%=CityArray(intIndex) %></option>
                      <% Next %>
                      </select></font></td>
                  </tr>
                </table>
                </td>
              </tr>
              <tr valign="bottom">
                <td width="163" bgcolor="#FFFFFF">
                &nbsp;</td>
                <td width="567" bgcolor="#FFFFFF" colspan="2">
                &nbsp;</td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td bgcolor="#FFFFFF">
                <img src="images/separator.gif" width="20" height="20"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="745" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" name="dates" id="dates">
              <tr valign="bottom">
                <td width="162">
                <img src="images/rental%20dates.gif" width="162" height="25"></td>
                <td width="132">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">Begin: 
                (mm/dd/yy)</font></td>
                <td width="88">
                <!--
                <input type="text" name="begin_date" class="fsmall" style="width:85" size="10" maxlength="10" onfocus="javascript:vDateType='1'" onkeyup="DateFormat(this,this.value,event,false,'1')" onblur="DateFormat(this,this.value,event,true,'1')">
                -->
				<% Dim datBeginDate_value
				   Dim datEndDate_value
				%>
				
                <% If adoRS4.Fields("exact_dates").Value = 0 Then %> 
                
                	<% If adoRS4.Fields("begin_arv_dt").Value < Now Then %>
                	<!--
                	<input name="begin_date" class="fsmall" value="<%=FormatDateTime(DateAdd("d", Now, DateDiff("d", adoRS4.Fields("inserted").Value, adoRS4.Fields("begin_arv_dt").Value)), 2) %>" maxlength="12" size="12" onblur="if(value=='')value='mm/dd/yy'" onfocus="if(value=='mm/dd/yy')value=''"></td>
                	-->
                	<% datBeginDate_value = FormatDateTime(DateAdd("d", Now, DateDiff("d", adoRS4.Fields("inserted").Value, adoRS4.Fields("begin_arv_dt").Value)), 2) %>
                
                	<% Else %>
                	<!--
                	<input name="begin_date" class="fsmall" value="<%=adoRS4.Fields("begin_arv_dt").Value %>" maxlength="12" size="12" onblur="if(value=='')value='mm/dd/yy'" onfocusin="if(value=='mm/dd/yy')value=''"></td>
                	-->
                	<% datBeginDate_value = adoRS4.Fields("begin_arv_dt").Value %>
                
                	<% End If %> 

                <% Else %>
                
					<% If adoRS4.Fields("begin_arv_dt").Value < Now Then %>
					<!--                	
                	<input name="begin_date" class="fsmall" value="<%=FormatDateTime(DateAdd("d", Now, 1), 2) %>" maxlength="12" size="12" onblur="if(value=='')value='mm/dd/yy'" onfocusin="if(value=='mm/dd/yy')value=''"></td>
                	-->
                	<% datBeginDate_value = FormatDateTime(DateAdd("d", Now, 1), 2) %>

	                <% Else %>
					<!--	
	                <input name="begin_date" class="fsmall" value="<%=adoRS4.Fields("begin_arv_dt").Value %>" maxlength="12" size="12" onblur="if(value=='')value='mm/dd/yy'" onfocusin="if(value=='mm/dd/yy')value=''"></td>
	                -->
	                <% dateBeginDate_value = adoRS4.Fields("begin_arv_dt").Value %>
	 	
		            <% End If %> 
		            
		        <% End If %>

				<div class="cbrow" id="fromrowed">
				<input type="text" name="begin_date" id="rental_begin_date" class="cb_txtdate" value="<%=datBeginDate_value%>" onfocus="openCal(this,'rental_begin_date','rental_end_date','calbox','fromrowed','us','vertical');if(this.value=='mm/dd/yyyy')this.value=''" title="mm/dd/yyyy" size="20" >
				</div>

               
			 <!-- Euro dd/mm/yyyy 
                <input type="text" name="begin_date" class="fsmall" style="width:85"  size='10' maxlength="10" onFocus="javascript:vDateType='3'" onKeyUp="DateFormat(this,this.value,event,false,'3')" onBlur="DateFormat(this,this.value,event,true,'3')"></td>
			  -->
			  	
                <td width="22" valign="middle" >
				&nbsp;
				<!--
                <img src="images/cal.gif" id="dtimg1" style="position:relative" border="0" title="View Calendar" width="16" height="16" onclick="ShowCalendar(document.search_criteria.dtimg1, document.search_criteria.begin_date, null, '<%=FormatDateTime(DateAdd("d", 1, Now),2) %>', '<%=FormatDateTime(DateAdd("d", 360, Now),2) %> - 1');event.cancelBubble=true;">
                -->
                <font color="#080000">


            </font>
                </td>
                <td width="21">
                <img src="images/separator.gif" width="20" height="27"> </td>
                <td width="124">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">End: 
                (mm/dd/yy)</font></td>
                <td width="82">
                <!-- 
              <input type="text" name="end_date" class="fsmall" style="width:85" size="10" maxlength="10" onfocus="javascript:vDateType='1'" onkeyup="DateFormat(this,this.value,event,false,'1')" onblur="DateFormat(this,this.value,event,true,'1')">
              -->
                <% If adoRS4.Fields("exact_dates").Value = 0 Then %> 
	                <% If adoRS4.Fields("end_arv_dt").Value < Now Then %>
						<!--
		                <input name="end_date" class="fsmall" value="<%=FormatDateTime(DateAdd("d", Now, DateDiff("d", adoRS4.Fields("inserted").Value, adoRS4.Fields("end_arv_dt").Value)), 2) %>" maxlength="12" size="12" onblur="verifyDateOrder();if(value=='')value='mm/dd/yy'"  onfocus="if(value=='mm/dd/yy')value=''"></td>
		                -->
		                <% datEndDate_value = FormatDateTime(DateAdd("d", Now, DateDiff("d", adoRS4.Fields("inserted").Value, adoRS4.Fields("end_arv_dt").Value)), 2) %>
	                <% Else %>
						<!--
		                <input name="end_date" class="fsmall" value="<%=adoRS4.Fields("end_arv_dt").Value %>" maxlength="12" size="12" onblur="verifyDateOrder();if(value=='')value='mm/dd/yy'" onfocus="if(value=='mm/dd/yy')value=''"></td>
		                -->
		                <% dateendDate_value = adoRS4.Fields("end_arv_dt").Value %>
	                <% End If %> 
	            <% Else %> 
	            	<% If adoRS4.Fields("end_arv_dt").Value < Now Then %>
	            		<!--
                		<input name="end_date" class="fsmall" value="<%=FormatDateTime(DateAdd("d", Now, 1), 2) %>" maxlength="12" size="12" onblur="verifyDateOrder();if(value=='')value='mm/dd/yy'" onfocus="if(value=='mm/dd/yy')value=''"></td>
                		-->
                		<% datEndDate_value = FormatDateTime(DateAdd("d", Now, 1), 2) %>
                	<% Else %>
                		<!--
                		<input name="end_date" class="fsmall" value="<%=adoRS4.Fields("end_arv_dt").Value %>" maxlength="12" size="12" onblur="verifyDateOrder();if(value=='')value='mm/dd/yy'" onfocus="if(value=='mm/dd/yy')value=''"></td>
                		-->
                		<% datEdnDate_value = adoRS4.Fields("end_arv_dt").Value %>
                	<% End If %> 
                <% End If %>

				<div class="cbrow" id="torowed">
					<input type="text" name="end_date" id="rental_end_date" class="cb_txtdate" value="<%=datEndDate_value%>" onfocus="openCal(this,'rental_begin_date','rental_end_date','calbox','torowed','us','vertical');if(this.value=='mm/dd/yyyy')this.value=''" title="mm/dd/yyyy" onclick="clearTimeout(t_calcloser)" size="20" >
				</div>	                
                
                </td>
                <td width="20">
                <!--
                <img src="images/cal.gif" id="dtimg2" style="position:relative" border="0" title="View Calendar" width="16" height="16" onclick="ShowCalendar(document.search_criteria.dtimg2, document.search_criteria.end_date, null, '<%=FormatDateTime(DateAdd("d", 1, Now),2) %>', '<%=FormatDateTime(DateAdd("d", 360, Now),2) %> - 1');event.cancelBubble=true;">
                -->
                </td>
                <td width="54">&nbsp;</td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="20" height="20"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom">
                <td width="162">
                <img src="images/separator.gif" width="162" height="20"></td>
                <td width="121">
                <font face="Vendana, Arial, Helvetica, sans-serif" size="2">ABP(Relative) 
                or Absolute:</font> </td>
                <td width="450">
                <select name="exact_dates" style="width:230" size="1">
                <% If adoRS4.Fields("exact_dates").Value = 0 Then 	%>
	                <option value="1">Date Absolute (those exact dates)</option>
	                <option selected value="0">Date Relative (# days from today)</option>
                <% Else If adoRS4.Fields("exact_dates").Value = 1 Then %>
	                <option selected value="1">Date Absolute (those exact dates)</option>
	                <option value="0">Date Relative (# days from today)</option>
                <% Else %>
	                <option value="1">Date Absolute (those exact dates)</option>
	                <option selected value="0">Date Relative (# days from today)</option>
                <% End If %>

                <% End If %>

                </select> <input type=button onclick=two() value="Object Offset" ></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="20" height="20"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom">
                <td width="164">
                <img title="Click to select all days of the week" src="images/ti_days_of_week.gif" width="162" height="25" border="0" onclick="check_all_days('dow_list')"></td>
                <td width="15">
                &nbsp;</td>
                <td width="25">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <% If Instr(adoRS4.Fields("dow_list").Value, "1") Then %>
                	<input type="checkbox" name="dow_list" value="1" checked> 
                <% Else %>
                	<input type="checkbox" name="dow_list" value="1"> 
                <% End If %>
                </font></td>
                <td width="27">
                <font size="2">Sun</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                </font></td>
                <td width="25">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif"><% If Instr(adoRS4.Fields("dow_list").Value, "2") Then %>
                <input type="checkbox" name="dow_list" value="2" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="2"> <% End If %>
                </font></td>
                <td width="30">
                <font size="2">Mon</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                </font></td>
                <td width="22">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif"><% If Instr(adoRS4.Fields("dow_list").Value, "3") Then %>
                <input type="checkbox" name="dow_list" value="3" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="3"> <% End If %>
                </font></td>
                <td width="26">
                <font size="2">Tue</font>
                </td>
                <td width="23">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif"><% If Instr(adoRS4.Fields("dow_list").Value, "4") Then %>
                <input type="checkbox" name="dow_list" value="4" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="4"> <% End If %>
                </font></td>
                <td width="31">
                <font size="2">Wed</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                </font></td>
                <td width="24">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif"><% If Instr(adoRS4.Fields("dow_list").Value, "5") Then %>
                <input type="checkbox" name="dow_list" value="5" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="5"> <% End If %>
                </font></td>
                <td width="24">
                <font size="2">Thu</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                </font></td>
                <td width="22">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif"><% If Instr(adoRS4.Fields("dow_list").Value, "6") Then %>
                <input type="checkbox" name="dow_list" value="6" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="6"> <% End If %>
                </font></td>
                <td width="17">
                <font size="2">Fri</font></td>
                <td width="210">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif"><% If Instr(adoRS4.Fields("dow_list").Value, "7") Then %>
                <input type="checkbox" name="dow_list" value="7" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="7"> <% End If %>
                Sat</font></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="20" height="20"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/separator.gif" width="20" height="15"></td>
              </tr>
            </table>
            <table width="748" border="0" cellpadding="2" cellspacing="0">
              <tr valign="bottom">
                <td width="162">&nbsp;</td>
                <td width="69">&nbsp;</td>
                <td width="54">&nbsp;</td>
                <td width="45">&nbsp;</td>
                <td width="62">&nbsp;</td>
                <td width="32">&nbsp;</td>
                <td width="296">
                <img src="images/separator.gif" width="20" height="27"> </td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/separator.gif" width="20" height="20"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="745" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom">
                <td width="163">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <img src="images/ti_details.gif" width="162" height="25"></font></td>
                <td width="97">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;Pick-up 
                Time :</font></td>
                <td width="92">
                <select name="arrival_time" style="width:90" size="1"> 
                <% Dim strTime
                   Dim strTimeValue 
                   Dim strCheckTime
                   Dim strTimeSelected
                   
                   strCheckTime = trim(adoRS4.Fields("arv_tm").Value)
                   

                %>
                	 <option value='24:00'>Midnight</option>
                <%

                   For intIndex = 1 To 11  
                     strTime = intIndex & ":00 am"
                     strTimeValue = intIndex & ":00"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%

                   Next   

				   If (strCheckTime = "12:00") Or IsNull(strCheckTime) Then	
                %>
                	 <option selected value='12:00'>Noon</option>
                <%
				   Else
                %>
                	 <option value='12:00'>Noon</option>
                <%
				   End If

                   For intIndex = 1 To 11  
                     strTime = intIndex & ":00 pm"
                     strTimeValue = intIndex + 12 & ":00"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                   Next   
                %>
                </select></td>
                <td width="99"><font size="2">Drop-off</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif"> 
                Time :</font> </td>
                <td width="130">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;<select name="return_time" style="width:90">
		        <% strCheckTime = trim(adoRS4.Fields("rtrn_tm").Value)

                %>
                	 <option value='24:00'>Midnight</option>
                <%

                   For intIndex = 1 To 11  
                     strTime = intIndex & ":00 am"
                     strTimeValue = intIndex & ":00"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%

                   Next   


				   If strCheckTime = "12:00" Or strCheckTime = "" Then	
                %>
                	 <option selected value='12:00'>Noon</option>
                <%
				   Else
                %>
                	 <option value='12:00'>Noon</option>
                <%
				   End If



                   For intIndex = 1 To 11  
                     strTime = intIndex & ":00 pm"
                     strTimeValue = intIndex + 12 & ":00"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                   Next   
                %>
                      </select></font></td>
                <td width="106">&nbsp;</td>
                <td width="45">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
              </tr>
              <tr valign="bottom">
                <td width="163">&nbsp;</td>
                <td width="97">&nbsp;</td>
                <td width="92">&nbsp;</td>
                <td width="99">&nbsp;</td>
                <td width="130">&nbsp;</td>
                <td width="106">&nbsp;</td>
                <td width="45">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td width="163">&nbsp;</td>
                <td width="97">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">Length 
                of Rent:</font></td>
                <td width="92">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <select name="lor" style="width:90; text-align:right" size="1"><% If adoRS4.Fields("lor").Value = 1 Then %>
                <option value='1' selected>1 day</option>
                <% Else %>
                <option value='1'>1 day</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 2 Then %>
                <option value='2' selected>2 days</option>
                <% Else %>
                <option value='2'>2 days</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 3 Then %>
                <option  value='3' selected>3 days</option>
                <% Else %>
                <option  value='3' >3 days</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 4 Then %>
                <option  value='4' selected>4 days</option>
                <% Else %>
                <option  value='4' >4 days</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 5 Then %>
                <option  value='5' selected>5 days</option>
                <% Else %>
                <option value='5'>5 days</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 6 Then %>
                <option  value='6' selected>6 days</option>
                <% Else %>
                <option  value='6'>6 days</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 7 Then %>
                <option  value='7' selected>7 days</option>
                <% Else %>
                <option  value='7'>7 days</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 8 Then %>
                <option  value='8' selected>8 days</option>
                <% Else %>
                <option  value='8'>8 days</option>
                <% End If %> <% If adoRS4.Fields("lor").Value = 9 Then  %>
                <option  value='9' selected>9 days</option>
                <% Else %>
                <option  value='9'>9 days</option>
                <% End If %> 
                <% If adoRS4.Fields("lor").Value = 10 Then %>
                <option  value='10' selected>10 days</option>
                <% Else %>
                <option  value='10'>10 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 11 Then %>
                <option  value='11' selected>10 days</option>
                <% Else %>
                <option  value='11'>11 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 12 Then %>
                <option  value='12' selected>10 days</option>
                <% Else %>
                <option  value='12'>12 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 13 Then %>
                <option  value='13' selected>13 days</option>
                <% Else %>
                <option  value='13'>13 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 14 Then %>
                <option  value='14' selected>14 days</option>
                <% Else %>
                <option  value='14'>14 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 15 Then %>
                <option  value='15' selected>15 days</option>
                <% Else %>
                <option  value='15'>15 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 16 Then %>
                <option  value='16' selected>16 days</option>
                <% Else %>
                <option  value='16'>16 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 17 Then %>
                <option  value='17' selected>17 days</option>
                <% Else %>
                <option  value='17'>17 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 18 Then %>
                <option  value='18' selected>18 days</option>
                <% Else %>
                <option  value='18'>18 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 19 Then %>
                <option  value='19' selected>19 days</option>
                <% Else %>
                <option  value='19'>19 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 20 Then %>
                <option  value='20' selected>20 days</option>
                <% Else %>
                <option  value='20'>20 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 21 Then %>
                <option  value='21' selected>21 days</option>
                <% Else %>
                <option  value='21'>21 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 22 Then %>
                <option  value='22' selected>22 days</option>
                <% Else %>
                <option  value='22'>22 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 23 Then %>
                <option  value='23' selected>23 days</option>
                <% Else %>
                <option  value='23'>23 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 24 Then %>
                <option  value='24' selected>24 days</option>
                <% Else %>
                <option  value='24'>24 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 25 Then %>
                <option  value='25' selected>25 days</option>
                <% Else %>
                <option  value='25'>25 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 26 Then %>
                <option  value='26' selected>26 days</option>
                <% Else %>
                <option  value='26'>26 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 27 Then %>
                <option  value='27' selected>27 days</option>
                <% Else %>
                <option  value='27'>27 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 28 Then %>
                <option  value='28' selected>28 days</option>
                <% Else %>
                <option  value='28'>28 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 29 Then %>
                <option  value='29' selected>29 days</option>
                <% Else %>
                <option  value='29'>29 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 30 Then %>
                <option  value='30' selected>30 days</option>
                <% Else %>
                <option  value='30'>30 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 31 Then %>
                <option  value='31' selected>31 days</option>
                <% Else %>
                <option  value='31'>31 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 32 Then %>
                <option  value='32' selected>32 days</option>
                <% Else %>
                <option  value='32'>32 days</option>
                <% End If %>

				<% If adoRS4.Fields("lor").Value = 1000 Then %>
                <option  value='1000' selected>Daily</option>
                <% Else %>
                <option  value='1000'>Daily</option>
                <% End If %>

				<% If adoRS4.Fields("lor").Value = 1001 Then %>
                <option  value='1001' selected>Wkend Day</option>
                <% Else %>
                <option  value='1001'>Wkend Day</option>
                <% End If %>


				<% If adoRS4.Fields("lor").Value = 1002 Then %>
                <option  value='1002' selected>Weekly</option>
                <% Else %>
                <option  value='1002'>Weekly</option>
                <% End If %>

                </select>
                </font>
                </td>
                <td width="99">
                <!-- 
              <font size="2" face="Vendana, Arial, Helvetica, sans-serif">Mileage 
              Charges :</font>
    			--></td>
                <td width="130">
                <!--              <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
              <select name="mileage_charges" style="width:90" size="1">
              <option selected value="0">Display</option>
              <option value="1">Disregard</option>
              </select></font>
--></td>
                <td width="106">&nbsp;</td>
                <td width="45">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td width="163">&nbsp;</td>
                <td width="97">&nbsp;</td>
                <td width="92">&nbsp;</td>
                <td width="99">&nbsp;</td>
                <td width="130">&nbsp;</td>
                <td width="106">&nbsp;</td>
                <td width="45">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td width="163">&nbsp;</td>
                <td width="97"><font size="2">Data Source</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">:</font>
                </td>
                <td width="306" colspan="3">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <select name="data_source" style="width:325; text-align:right">
                <% While adoRS1.EOF = False %>
                <% If adoRS4.Fields("data_sources").Value = adoRS1.Fields("data_source").Value Then %>
                <option selected value="<%=adoRS1.Fields("data_source").Value %>">
                <%=adoRS1.Fields("name").Value %></option>
                <% ElseIf (adoRS4.Fields("data_sources").Value = "") And (adoRS1.Fields("data_source").Value = "ORB") Then %>
                <option selected value="<%=adoRS1.Fields("data_source").Value %>">
                <%=adoRS1.Fields("name").Value %></option>
                <% Else %>
                <option value="<%=adoRS1.Fields("data_source").Value %>"><%=adoRS1.Fields("name").Value %>
                </option>
                <% End If %> <% 
					 adoRS1.MoveNext
				   Wend
				   Set adoRS1 = Nothing
				%></select> </font></td>
                <td width="106">&nbsp;</td>
                <td width="45">&nbsp;</td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">&nbsp;</td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/separator.gif" width="20" height="15"></td>
              </tr>
            </table>
            <table width="676" border="0" cellpadding="0" cellspacing="0">
              <tr valign="bottom">
                <td width="162" valign="top">
                <img src="images/ti_car_types.gif" width="162" height="50"></td>
                <td width="20" valign="top">
                <img src="images/separator.gif" width="20" height="20"></td>
                <td width="483" valign="bottom">
                <table width="473" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
                  <tr>
                    <td width="473" colspan="3">
                    <div align="center">
                      <font size="2">(use buttons to move between selected/unselected)</font></div>
                    <div align="center">
                      <font size="2">Unselected types&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      Selected types&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
                    </div>
                    </td>
                  </tr>
                  <tr>
                    <td width="211">
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                      <select name="unselected_car_types" size="5" style="width:200;"  multiple>
                      <% Dim CarTypes                                        %> 
                      <% CarTypes = adoRS4.Fields("shop_car_type_cds").Value %>
                      <% While adoRS5.EOF = False                            %> 
                      <% If (InStr(CarTypes, adoRS5.Fields("car_type_cd").Value) = 0) Or (CarTypes = "") Then %>
                      <option value="<%=adoRS5.Fields("car_type_cd").Value %>"><%=adoRS5.Fields("car_type_cd").Value %>
                      </option>
                      <% End If %> 
                      <% adoRS5.MoveNext %> 
                      <% Wend %>
                      <!--                    
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
--></select></font></td>
                    <td width="31">
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_car_types, document.search_criteria.selected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_right.GIF" width="23" height="22" longdesc="Move right"    ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_car_types, document.search_criteria.selected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_right_all.GIF" width="23" height="22"    ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_car_types, document.search_criteria.unselected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_left.GIF" width="23" height="22"    ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_car_types, document.search_criteria.unselected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_left_all.GIF" width="23" height="22"    ></a></font></td>                    
                    <td width="231">
                    <font color="#080000">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="selected_car_types" size="5" style="width:200;" multiple>
                      <% Dim CarTypeArray                                    %> <% CarTypeArray = Split(CarTypes, ",")                 %>
                      <% For intIndex = 0 To UBound(CarTypeArray)            %>
                      <option value="<%=CarTypeArray(intIndex) %>"><%=CarTypeArray(intIndex) %>
                      </option>
                      <% Next %></select></font></font></td>
                  </tr>
                </table>
                </td>
              </tr>
            </table>
            <table width="677" border="0" cellpadding="0" cellspacing="0">
              <tr valign="bottom">
                <td width="162" valign="top">
                <img src="images/separator.gif" width="162" height="20"></td>
                <td width="20" valign="top">
                <img src="images/separator.gif" width="20" height="20"></td>
                <td width="485" valign="bottom">
                <table width="475" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
                  <tr>
                    <td width="475" colspan="3">
                    <div align="center">
                      <div align="center">
&nbsp;</div>
                      <div align="center">
                        <font size="2">(use buttons to move between selected/unselected)</font></div>
                      <div align="center">
                        <font size="2">Unselected companies&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        Selected companies&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; </font>
                      </div>
                    </div>
                    </td>
                  </tr>
                  <tr>
                    <td width="209">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="unselected_companies" size="5" style="width:200;" multiple>
                      <% Dim Companies                                       %> 
                      <% Dim SelectedCompanies							     %>
                      <% Companies = adoRS4.Fields("vend_cds").Value         %> 
                      <% While adoRS2.EOF = False                            %>
                      <% If ((InStr(Companies, adoRS2.Fields("vendor_cd").Value) = 0) Or (Companies = "")) And (adoRS2.Fields("vendor_cd").Value <> strSelfVendCd) Then %>
                      	<option value="<%=adoRS2.Fields("vendor_cd").Value %>"><%=adoRS2.Fields("vendor_name").Value %></option>
                      <% Else 		                    
		                    SelectedCompanies = SelectedCompanies & TRIM(adoRS2.Fields("vendor_cd").Value) & "," & adoRS2.Fields("vendor_name").Value & "," 
						
						   End If 
						   adoRS2.MoveNext 
					   Wend 
					   Set adoRS2 = Nothing %>
					</select>
					</font>
					</td>
                    <td width="31">
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_companies, document.search_criteria.selected_companies, false );update_company_list();return false;">
                      <img border="0" src="images/move_right.GIF" width="23" height="22"    ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_companies, document.search_criteria.selected_companies, true );update_company_list();return false;">
                      <img border="0" src="images/move_right_all.GIF" width="23" height="22"    ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_companies, document.search_criteria.unselected_companies, false );update_company_list();return false;">
                      <img border="0" src="images/move_left.GIF" width="23" height="22"    ></a></font><br>
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
    
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_companies, document.search_criteria.unselected_companies, true );update_company_list();return false;">
                      <img border="0" src="images/move_left_all.GIF" width="23" height="22"    ></a></font></td>
                    <td width="228">
                      <font size="2" face="Vendana, Arial, Helvetica, sans-serif" color="#080000">
                      <select name="selected_companies" size="5" style="width:200;" multiple>
                      <% Dim CompanyArray                                            %> 
                      <% CompanyArray = Split(SelectedCompanies, ",")                %>
                      <% For intIndex = 0 To UBound(CompanyArray) - 1 Step 2         %>
                        <option value="<%=CompanyArray(intIndex) %>"><%=CompanyArray(intIndex + 1) %></option>
                      <% Next %>
                      </select>
                      </font>
                    </td>
                  </tr>
                </table>
                </td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/separator.gif" width="20" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/separator.gif" width="20" height="20"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="745" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom">
                <td width="162">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <img src="images/ti_display_options.gif" width="162" height="25"></font></td>
                <td width="119">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">Highlighted 
                Company : </font></td>
                <td width="144">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <select name="highlighted_company" style="width:200" size="1">
                <% If Len(SelectedCompanies) = 0 Then                          %>
	                <option selected>companies must be selected</option>
				<% Else														   %>
                <option >companies must be selected</option>
                <% For intIndex = 0 To UBound(CompanyArray) - 1 Step 2         %>
                	<% If (adoRS4.Fields("highlight_vendor").Value = CompanyArray(intIndex)) Then %>
		                <option selected value="<%=CompanyArray(intIndex) %>"><%=CompanyArray(intIndex + 1) %></option>
					<% Else %>
		                <option value="<%=CompanyArray(intIndex) %>"><%=CompanyArray(intIndex + 1) %></option>
					<% End If %>
                <% Next %>
                <% End If                                                      %>
                </select></font></td>
                <td width="20">
                <!--               
              <input type="checkbox" name="rate_drilldown" value="enable">
--></td>
                <td width="280">
                <!--              
              <font size="2" face="Vendana, Arial, Helvetica, sans-serif">Enable 
              Rate Drill Down</font> 
--></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                  <tr valign="bottom">
                    <td width="161">
                    <img src="images/separator.gif" width="30" height="15"></td>
                  </tr>
                </table>
                <table width="734" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                  <tr valign="bottom">
                    <td width="162">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <img src="images/separator.gif" width="162" height="20"></font></td>
                    <td width="119">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Rates Shown </font></td>
                    <td width="146">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="rates_shown" style="width:125" disabled >
                    <option value="1">Rate/Est TotChg</option>
                    <option selected value="2">Rate Only</option>
                    <option value="3">Total Charge only</option>
                    </select> </font></td>
                    <td width="97">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Currency :</font> </td>
                    <td width="190">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="display_currency" style="width:125" size="1">
                    <option value="USD" selected>USD</option>
                    <option value="CAD">CAD</option>
                    <option value="GBP">GBP</option>
                    <option value="EUR">EUR</option>
                    </select> </font></td>
                  </tr>
                </table>
                </td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom">
                <td width="161">
                <img src="images/separator.gif" width="30" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/separator.gif" width="20" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0">
              <tr valign="bottom">
                <td width="162">
                <img src="images/ti_rate_changes.gif" width="162" height="25"></td>
                <td width="575">
                <table width="578" border="0" cellpadding="2" cellspacing="0">
                  <tr valign="bottom">
                    <td width="190">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Days Prior to Compare Rates:</font> </td>
                    <td width="76">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="days_prior" size="1" disabled >
                    <option selected value="0">None</option>
                    <option value="1">1</option>
                    <option value="2">2</option>
                    <option value="3">3</option>
                    <option value="4">4</option>
                    <option value="5">5</option>
                    <option value="6">6</option>
                    <option value="7">7</option>
                    <option value="8">8</option>
                    <option value="9">9</option>
                    <option value="10">10</option>
                    </select> </font></td>
                    <td width="117">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Charge Detail :</font> </td>
                    <td width="179">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="change_detail" size="1" disabled >
                    <option selected value="1">Previous Rate and Direction</option>
                    <option value="2">Direction</option>
                    <option value="3">Previous Rate</option>
                    </select> </font></td>
                  </tr>
                </table>
                </td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/separator.gif" width="20" height="20"></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <p><img src="images/ti_batch_options.gif" width="365" height="24"></p>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><img src="images/ruler.gif" width="745" height="2"></td>
              </tr>
            </table>
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="745" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
              <tr valign="bottom">
                <td width="162">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                <img src="images/ti_actions.gif" width="162" height="25"></font></td>
                <td width="148">&nbsp;</td>
                <td width="423">
                <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
              </tr>
            </table>
            <table width="663" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif" width="799">
                <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                  <tr valign="bottom">
                    <td width="161">
                    <table width="546" border="0" cellpadding="2" cellspacing="0">
                      <tr valign="bottom">
                        <td width="162">
                        <img src="images/separator.gif" width="162" height="20"></td>
                        <td width="184">
                        <p>
                        <input type="radio" value="html" checked name="output"><font size="2" face="Vendana, Arial, Helvetica, sans-serif">HTML 
                        Presentation</font></p>
                        </td>
                        <td width="188">
                        <select name="html_style" style="width:175" size="1" disabled>
                        <option value="2">By Date</option>
                        <option value="3">By Car Type</option>
                        <option selected value="1">By Date and Car Type</option>
                        <option value="4">Lowest Rate Detail</option>
                        <option value="5">All Rate Detail</option>
                        </select></td>
                      </tr>
                    </table>
                    <table width="546" border="0" cellpadding="2" cellspacing="0">
                      <tr valign="bottom">
                        <td width="162">
                        <img src="images/separator.gif" width="162" height="20"></td>
                        <td width="184">
                        <input type="radio" value="ftp" name="output" ><font size="2">Save 
                        to FTP server</font></td>
                        <td width="188">
                        <select name="ftp_style" style="width:175" size="1">
                        <option selected value="1">Default layout</option>
                        </select></td>
                      </tr>
                    </table>
                    <table width="546" border="0" cellpadding="2" cellspacing="0">
                      <tr valign="bottom">
                        <td width="162">
                        <img src="images/separator.gif" width="162" height="20"></td>
                        <td width="184">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                        <input type="radio" value="xls" name="output" disabled >Excel XLS 
                        Presentation</font></td>
                        <td width="188">
                        <select name="xls_type" style="width:175" disabled size="1" >
                        <option selected value="1">No formats installed</option>
                        <option value="Custom">Custom</option>
                        </select></td>
                      </tr>
                    </table>
                    </td>
                  </tr>
                </table>
                <table width="717" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                  <tr valign="bottom">
                    <td width="162">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <img src="images/separator.gif" width="162" height="20"></font></td>
                    <td width="158">&nbsp;</td>
                    <td width="59">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                    <td width="106">&nbsp;</td>
                    <td width="212">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  </tr>
                </table>
                <table width="712" border="0" cellpadding="2" cellspacing="0">
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td>
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Action : </font></td>
                    <td>
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="search_action" style="width:175" size="1">
                    <% If adoRS4.Fields("action_id").Value = 1 Then %>
                    <option selected value="1">Search</option>
					<% Else %>
                    <option value="1">Search</option>
                    <% End If %>
					
                    <% If adoRS4.Fields("action_id").Value = 2 Then %>
	                    <option selected value="2">Search & email all</option>
					<% Else %>
	                    <option value="2">Search & email all</option>
                    <% End If %>

                    <% If adoRS4.Fields("action_id").Value = 3 Then %>
	                    <option selected value="3">Search & email alerts</option>
					<% Else %>
	                    <option value="3">Search & email alerts</option>
                    <% End If %>
					
					
                    </select></font></td>
                  </tr>
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td>
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp; 
                    E-mail Address: </font></td>
                    <td>
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <input type="text" name="recipient_address" size="26" style="width:175" value="<%=adoRS4.Fields("email_address").Value %>"></font></td>
                  </tr>
                  <tr valign="bottom">
                    <td><font size="2">
                    <img src="images/separator.gif" width="162" height="20"></font></td>
                    <td>
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Shop POS:</font></td>
                    <td><select name="pos" style="width:175">
                    <option value="GB">GB - United Kingdom</option>
                    <option selected value="US">US - United States</option>
                    <option value="FR">FR - France</option>
                    <option value="CA">CA - Canada</option>
                    </select></td>
                  </tr>
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td>
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Airline Arrival Required</font></td>
                    <td><input type="checkbox" name="airline" value="required"></td>
                  </tr>
                  <tr valign="bottom">
                    <td width="162">
                    <img src="images/separator.gif" width="162" height="20"></td>
                    <td width="184">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Discount Code </font></td>
                    <td width="354">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <input type="text" name="discount_code" size="20" style="width:175">
                    </font><select name="discount_car_company">
                    <option selected>Alamo</option>
                    <option>Avis</option>
                    <option>Hertz</option>
                    <option>Thrifity</option>
                    </select></td>
                  </tr>
                </table>
                <table width="713" border="0" cellpadding="2" cellspacing="0">
                  <tr valign="bottom">
                    <td width="162">
                    <img src="images/separator.gif" width="162" height="20"></td>
                    <td width="184">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    Scheduled Action Time: </font></td>
                    <td width="355">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                    <select name="scheduled_time" style="width:175" size="1">
                <% If IsNull(adoRS4.Fields("scheduled_dttm").Value) Then 
                	strCheckTime ="" 
                %>
                	<option selected value =''>Perform Search Now</option>    
                
		        <% Else %>
                	<option value =''>Perform Search Now</option>    
				<%
		            strCheckTime = trim(FormatDateTime(adoRS4.Fields("scheduled_dttm").Value, 4))
		            
		           End If 

                %>
                
                <% If InStr(1, strCheckTime, "00:00") > 0 Then %>
                	 <option selected value='00:00'>Midnight</option>
                
                <% Else %>
                	 <option value='00:00'>Midnight</option>
                
                <% End If %>
                
                <%

                   For intIndex = 1 To 11  
                     strTime = intIndex & ":00 am"
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%

                   Next   


				   If strCheckTime = "12:00" Then	
                %>
                	 <option selected value='12:00'>Noon</option>
                <%
				   Else
                %>
                	 <option value='12:00'>Noon</option>
                <%
				   End If



                   For intIndex = 1 To 11  
                     strTime = intIndex & ":00 pm"
                     strTimeValue = trim(FormatDateTime(strTime, 4))
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                   Next   
                %>
                    </select> (PST/PDT).</font></td>
                  </tr>
                </table>
                <table width="743" border="0" cellpadding="2" cellspacing="0">
                  <tr valign="bottom">
                    <td width="162">
                    <img src="images/separator.gif" width="162" height="20"></td>
                    <td width="184">
                    <font size="2" face="Vendana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp; 
                    Days of Week:</font></td>
                    <td width="385">
                    <table width="381" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                      <tr valign="bottom">
                        <td width="20">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "1") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="1" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="1"> 
		                <% End If %>
                        </font></td>
                        <td width="29">
                        <font size="2">Sun</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                        </font></td>
                        <td width="20">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "2") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="2" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="2"> 
		                <% End If %>
                        </font></td>
                        <td width="25">
                        <font size="2">Mon</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                        </font></td>
                        <td width="20">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "3") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="3" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="3"> 
		                <% End If %>
                        </font></td>
                        <td width="31">
                        <font size="2">Tue</font> </td>
                        <td width="20">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "4") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="4" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="4"> 
		                <% End If %>
                        </font></td>
                        <td width="25">
                        <font size="2">Wed</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                        </font></td>
                        <td width="20">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "5") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="5" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="5"> 
		                <% End If %>
                        </font></td>
                        <td width="19">
                        <font size="2">Thu</font><font size="2" face="Vendana, Arial, Helvetica, sans-serif">
                        </font></td>
                        <td width="20">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "6") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="6" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="6"> 
		                <% End If %>
                       </font></td>
                        <td width="20">
                        <font size="2">Fri</font></td>
                        <td width="20">
                        <font size="2" face="Vendana, Arial, Helvetica, sans-serif">
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "7") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="7" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="7"> 
		                <% End If %>
                       </font></td>
                        <td width="20">
                        <font size="2" color="#080000">Sat</font></td>
                      </tr>
                    </table>
                    </td>
                  </tr>
                  <tr valign="bottom">
                    <td width="162">&nbsp;</td>
                    <td width="184">&nbsp;</td>
                    <td width="385">&nbsp;</td>
                  </tr>
                  <tr valign="bottom">
                    <td width="162">&nbsp;</td>
                    <td width="569" colspan="2">
                    <input name="submit" type="submit" id="submit_request" value="Submit Request" class="rh_button" onclick="return confirmSubmit('searchnow')"></td>
                  </tr>
                  <tr valign="bottom">
                    <td width="162"></td>
                    <td width="569" colspan="2">&nbsp;</td>
                  </tr>
                </table>
                </td>
              </tr>
            </table>
            </font></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
  
   
              <p align="center">Debug section:<br>
              <input type="text" name="cities_list"  value="<%=adoRS4.Fields("city_cd").Value%>">
              <input type="text" name="car_type_list" value="<%=adoRS4.Fields("shop_car_type_cds").Value %>">
              <input type="text" name="company_list" value="<%=adoRS4.Fields("vend_cds").Value %>"> 
              
              <br>
              <input type="checkbox" name="debug" value="true" checked>Debug 
              (don't alter the database, just display the results)<br>
              strCheckTime=<%=strCheckTime  %><br>

</form>

			  <TEXTAREA id="report" rows=10 cols=50 wrap=physical style="margin-top:10px; border:1px solid #cccccc; font-family:arial; width:100%; height:50%" ></TEXTAREA>
              
              </p>
<form method="POST" action="" name="pos_values">
  <p><input type="text" name="x" id="x" size="20"></p>
  <p><input type="text" name="y" id="y"   size="20"></p>
</form>
<p>&nbsp;</p>
<p>&nbsp;</p>
       
           
	
              
              
<!--
<script language="JavaScript">
-->
		<!-- // create calendar object(s) just after form tag closed
			 // specify form element as the only parameter (document.forms['formname'].elements['inputname']);
			 // note: you can have as many calendar objects as you need for your application
			var cal6 = new calendar2(document.forms['search_criteria'].elements['begin_date']);
			cal6.year_scroll = false;
			cal6.time_comp = false;
			var cal7 = new calendar2(document.forms['search_criteria'].elements['end_date']);
			cal7.year_scroll = false;
			cal7.time_comp = false;
		//-->
<!--
		</script>
-->
<div id="calbox" class="calboxoff"></div>	
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<%	
	Rem Clean up
	Set adoCmd = Nothing 
	Set adoCmd1 = Nothing 
	Set adoCmd2 = Nothing 
	Set adoCmd3 = Nothing 
	Set adoCmd4 = Nothing
	Set adoCmd5 = Nothing 	
	
%>