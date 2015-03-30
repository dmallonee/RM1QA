<%@ Language=VBScript %>
<%
'Revisions
'When     Who What
'======== === ==========================================================
'
'
%>
<!-- #INCLUDE FILE="inc/login_check_ex.asp" -->
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180

	Dim strConn	
	Dim adoCmd	
	Dim adoRS
	Dim adoCmd2	
	Dim adoRS2
	Dim blnAllowSaving
    Dim strSelectedCar
    Dim strSelectedComp
	
	'Images for Page Header
	Dim rhBanner
	Dim pageLabelURL
	rhBanner = "images/top.jpg"
	pageLabelURL = "images/h_search_criteria.gif"

	Dim adoPrices
	Dim strUserId
	Dim strHighlightVendor
	
	On Error Resume Next
	
	strUserId     = Session("user_id")
	strSelfVendCd = Session("vend_cd")
	intRptLimit   = Session("rpt_limit")

    strProfileDesc = Request("profile_desc")
	strProfileCarType = Request("profile_car_type")
	strProfileCarCo = Request("profile_car_co")

  	Session("user_id") = strUserID
	strConn = Session("pro_con")

    if Request("submitform") <> "" then
        strSelectedCar = Replace(Request("selected_car_types")," ","")
        strSelectedComp = Replace(Request("selected_companies")," " ,"")
        strSelectedProf = Replace(Request("profile_id")," ","")
        strHighlightedComp = Replace(Request("highlighted_company")," ","")

	    Set adoCmd = CreateObject("ADODB.Command")
	    adoCmd.ActiveConnection =  strConn
	    adoCmd.CommandText = "car_shop_profile_bulk_update"
	    adoCmd.CommandType = 4
    	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_profile_ids", 200, 1,1024)
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 1024)
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_vend_cds", 200, 1, 1024)
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@highlight_vendor", 200, 1, 10)
		adoCmd.Parameters("@shop_profile_ids").Value = Trim(strSelectedProf) 
		adoCmd.Parameters("@shop_car_type_cds").Value = Trim(strSelectedCar) 
		adoCmd.Parameters("@shop_vend_cds").Value = Trim(strSelectedComp) 
		adoCmd.Parameters("@highlight_vendor").Value = Trim(strHighlightedComp) 
        Set adoRS = adoCmd.Execute
        set adoRS = nothing
        set adoCmd = nothing
    end if

    Rem Get the profiles
	Set adoCmd = CreateObject("ADODB.Command")
	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_profile_select"
	adoCmd.CommandType = 4
	adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 1024)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@enabled", 3, 1, 0)
 
    If Trim(Replace(strProfileDesc,"*","%")) <> "" Then
		adoCmd.Parameters("@desc").Value = Trim(Replace(strProfileDesc,"*","%")) 
	Else
		adoCmd.Parameters("@desc").Value = Null
	End If

	If Trim(Replace(strProfileCarType,"*","%")) <> "" Then
		adoCmd.Parameters("@shop_car_type_cds").Value = Trim(Replace(strProfileCarType,"*","%")) 
	Else
		adoCmd.Parameters("@shop_car_type_cds").Value = Null
	End If

	If Trim(Replace(strProfileCarCo,"*","%")) <> "" Then
		adoCmd.Parameters("@vend_cds").Value = Trim(Replace(strProfileCarCo,"*","%")) 
	Else
		adoCmd.Parameters("@vend_cds").Value = Null 
	End If
    adoCmd.Parameters("@user_id").Value = strUserId
    adoCmd.Parameters("@enabled").Value = 1

    Set adoRS = adoCmd.Execute

	Rem Get the vendors
	Set adoCmd2 = CreateObject("ADODB.Command")
	adoCmd2.ActiveConnection =  strConn
	adoCmd2.CommandText = "vendor_select"
	adoCmd2.CommandType = 4
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_cd",   200, 1,  2, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_name", 200, 1, 50, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@user_id",       3, 1,  0, strUserId)
	Set adoRS2 = adoCmd2.Execute

	Rem Get the car types
	Set adoCmd5 = CreateObject("ADODB.Command")
	adoCmd5.ActiveConnection =  strConn
	adoCmd5.CommandText = "car_type_select"
	adoCmd5.CommandType = 4
	adoCmd5.Parameters.Append adoCmd5.CreateParameter("@user_id", 3, 1, 0, strUserId)
	Set adoRS5 = adoCmd5.Execute
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>

<head>
<META http-equiv="Content-Language" content="en-us">
<META HTTP-EQUIV="refresh" CONTENT="900;URL=default_session.asp">
<title>Rate-Monitor by Rate-Highway, Inc. | Bulk Profile Updater</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script type="text/javascript" language="JavaScript" src="inc/ts_picker.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/center.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/date_check.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/multiple_select_support.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/multiple_select_support2.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/selectbox.js"></script>
<script type="text/javascript" language="Javascript">


    function update_company_list(reset_highlight, discount_list) {

        var list = new String("")
        var discount = discount_list.split('|')
        // the profile save saves them in reverse order, so we need to flip them back into order
        discount.reverse();


        if (reset_highlight == true) {
            removeAllOptions(document.getElementById('search_criteria').highlighted_company);
            addOption(document.getElementById('search_criteria').highlighted_company, "(please select one)", "", true);
        }
        // Clear out the discount codes
        //	deleteRows();
        var numoptions = document.getElementById('selected_companies').options.length;
        for (i = 0; i < numoptions; i++) {
            if (reset_highlight == true) {
                addOption(document.getElementById('highlighted_company'), document.getElementById('selected_companies').options[i].text, document.getElementById('selected_companies').options[i].value, false);
            }

            if (i == 0) {
                list = document.getElementById('selected_companies').options[i].value;
            }
            else {
                list = list + ',' + document.getElementById('selected_companies').options[i].value;
            }
        }

        document.getElementById('company_list').value = list;
        //document.getElementById('search_criteria').selected_companies.options.text; 

    }
    function update_car_type_list() {

        var list = new String("")

        for (i = 0; i < document.getElementById('search_criteria').selected_car_types.options.length; i++) {
            if (i == 0) {
                list = document.getElementById('search_criteria').selected_car_types.options[i].value
            }
            else {
                list = list + ',' + document.getElementById('search_criteria').selected_car_types.options[i].value
            }
        }

        document.getElementById('search_criteria').car_type_list.value = list
        //document.getElementById('search_criteria').selected_companies.options.text; 

    }

    function deleteRows() {
        // Check all rows in table. Length of table will change during loop if deletes occur.
        var tbl = document.getElementById('discount_codes');
        for (i = tbl.rows.length - 1; i != 0; i--) {
            tbl.deleteRow(i);
        } // End process all rows
        return true;
    } // End deleteRows function


    function validateForm() {
        var i;
        var obj;
        var count = 0;
        var why = "";
        for (i = 0; i < document.getElementById('search_criteria').elements.length; i++) {
            obj = document.getElementById('search_criteria').elements[i];
            if (obj.type == "checkbox" && obj.checked) {
                count++;
            }
        }
        var proflength = count;
        var typelength = document.getElementById('selected_car_types').options.length;
        var complength = document.getElementById('selected_companies').options.length;
        if (proflength == 0) {
            why += "Please select at least one profile to update.\n";
        }
        if (typelength == 0 && complength == 0) {
            why += "Please select at least one car type or company.\n";
        }
        if (complength > 0 && document.getElementById('highlighted_company').options[document.getElementById('highlighted_company').selectedIndex].value == "") {
            why += "You must select a comparison company\n";
        }
        if (why != "") {
            alert(why);
            return false;
        }
        for (i = 0; i < typelength; i++) {
            document.getElementById('selected_car_types').options[i].selected = true;
        }
        for (i = 0; i < complength; i++) {
            document.getElementById('selected_companies').options[i].selected = true;
        }
        return true;
    }

    function checkAll(field) {
        for (i = 0; i < field.length; i++) {
            field[i].checked = true;
        }
    }

    function uncheckAll(field) {

        for (i = 0; i < field.length; i++)
            field[i].checked = false;
    }
</script> 


<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css" >
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type="text/javascript" language="javascript" src="inc/header_menu_support.js" ></script>
<script language="javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/css_calendar_v2.js"></script>	
<style type="text/css">
.style1 {
	border-left-style: solid;
	border-left-width: 1px;
	border-right: 1px solid #C0C0C0;
	border-top-style: solid;
	border-top-width: 1px;
	border-bottom: 1px solid #C0C0C0;
}
.style2 {
	color: #FF0000;
}
td {
	font-size: x-small;
	font-weight: normal;
	color: #000000;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
body {
	top: 0px;
	right: 0px;
	bottom: 0px;
	left: 0px;
	background-color: #FFFFFF;
}
</style>
</head>

<body>
<a name="top"></a>

<div id="content"> 


<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_middle.jpg">
    <img src="images/top_left.jpg" width="423" height="91" alt=""></td>
    <td background="images/top_middle.jpg"></td>
    <td background="images/top_tile.gif">
    <img src="images/top_right.jpg" width="365" height="91" alt=""></td>
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
    <img src="images/med_bar.gif" width="12" height="8" alt=""></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/user_left.gif" width="580" height="31" alt=""></td>
        <td background="images/user_tile.gif">
        <table width="100" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td valign="bottom">
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>
                <div align="right">
                  <font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%=Request.Cookies("rate-monitor.com")("user_name")%></font></div>
                </td>
              </tr>
              <tr>
                <td><img src="images/separator.gif" width="183" height="6" alt=""></td>
              </tr>
            </table>
            </td>
            <td><img src="images/user_tile.gif" width="7" height="31" alt=""></td>
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
        <td><img src="images/h_search_profiles.gif" width="368" height="31" alt=""></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0" alt=""></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<h2 style="text-align:center;font:bold 18pt Arial">RATE HIGHWAY BULK PROFILE UPDATER (BETA)</h2>
<form method="POST" action="bulk_update_profiles.asp" name="search_profiles" class="search">
  <table align="center" border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="table2" background="images/alt_color.gif">
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182"></td>
      <td width="910" height="18" colspan="3">
      <font size="2" >To screen your profiles, enter any portion of the description.<br />You may use asterisks (*) as wild cards.<br />You may also screen for car types and companies.</font></td>
      <td align="right"><!--help button would go here--></td>
    </tr>
    <tr>
      <td width="19" height="26">&nbsp;</td>
      <td width="182"><img border="0" src="images/search.GIF" alt="Search"></td>
      <td width="134" height="26"><font size="2">Profile Description:</font></td>
      <td width="168" height="26">
      <input type="text" name="profile_desc" size="20" value="<%=strProfileDesc %>" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');"></td>
      <td width="608" height="26">
      <font size="2">
      <input type="submit" value="  Screen Profiles  " name="submit0" class="rh_button"></font></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="134" height="22">
      <font size="2">Car Type Code:</font></td>
      <td width="168" height="22">
      <input type="text" name="profile_car_type" size="20" value="<%=strProfileCarType %>" onfocus="this.className='focus';cl(this,'<%=strCarType %>');" onblur="this.className='';fl(this,'<%=strCarType %>');"></td>
      <td width="608" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="134" height="22">
      <font size="2">Car Company:</font></td>
      <td width="168" height="22">
      <input type="text" name="profile_car_co" size="20" value="<%=strProfileCarCo  %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
      <td width="608" height="22">&nbsp;</td>
    </tr>

    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="134" height="18">
      &nbsp;</td>
      <td width="168" height="18">
      &nbsp;</td>
      <td width="608" height="18">&nbsp;</td>
    </tr>
  </table>
 </form>

     <table align="center" border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4" id="table1">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
 </table>

<form method="POST" name="search_criteria" id="search_criteria" action="bulk_update_profiles.asp" onsubmit="return validateForm()">
            <table align="center" border="0" width="1110: cellpadding="2" style="border-collapse: collapse"  bordercolor="#FFFFFF" background="images/alt_color.gif">
            <tr>
                <td>
                    <table align="center">
                     <tr valign="bottom">
                    <td colspan="9">
                    <div align="center">
                      (use buttons to move between selected/unselected)
                    </div>
                    </td>
                  </tr>
                    <tr>
                      <td width="30"></td>
                    <td width="175px" align="center">Unselected car types</td>
                    <td width="30"></td>
                    <td width="175px" align="center">Selected car types</td>
                    <td width="24"></td>
                    <td width="175px"></td>
                     <td width="175px" align="center">Unselected companies</td>
                    <td width="24"></td>
                   <td width="175px" align="center">Selected companies</td>
                    <td width="24"></td>
                    <td width="150"></td>
                    <td width="175px" align="center">Comparison company</td>
                      <td width="30"></td>
                 </tr>
                  <tr>
                      <td width="30"></td>
                    <td align="right">
                      
                      <select name="unselected_car_types" size="12" style="width:175px"  multiple>
                      <% Dim CarTypes                                        %> 
                      <% Dim CarClasses                                      %> 
                      <% While adoRS5.EOF = False                            %> 
                      <option value="<%=adoRS5.Fields("car_type_cd").Value %>"><%=adoRS5.Fields("class_desc").Value %>
                      </option>
                      <% adoRS5.MoveNext %> 
                      <% Wend %>
					  </select></td>
                    <td>
                      
                    <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').unselected_car_types, document.getElementById('search_criteria').selected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_right.GIF" width="24" height="22" alt="Add the highlighted car types"    ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').unselected_car_types, document.getElementById('search_criteria').selected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_right_all.GIF" width="24" height="22"  alt="All all the car types"  ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').selected_car_types, document.getElementById('search_criteria').unselected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_left.GIF" width="24" height="22"  alt="Remove the highlighted car types"  ></a><br>
                      
    
                    <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').selected_car_types, document.getElementById('search_criteria').unselected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_left_all.GIF" width="24" height="22"  alt="Remove all the car types"  ></a></td>                    
                      <td>
                    <font color="#080000">
                    
                    <select name="selected_car_types" id="selected_car_types" size="12" style="width:175px" multiple>
                      <% Dim CarTypeArray                                    %> 
                      <% CarTypeArray = Split(CarTypes, ",")                 %>
                      <% For intIndex = 0 To UBound(CarTypeArray)            %>
                      <option value="<%=CarTypeArray(intIndex) %>"><%=CarTypeArray(intIndex) %>
                      </option>
                      <% Next %></select></font></td>
                    <td align="left">
                    <a href="javascript:void(0)" onclick="moveOptionUp(document.getElementById('search_criteria').selected_car_types);update_car_type_list();return false;">
                      <img border="0" src="images/up_button.GIF" width="24" height="24"  alt="Click this up arrow to modify the display order of the car types" ></a><br>
                    <font color="#080000">
                    <a href="javascript:void(0)" onclick="moveOptionDown(document.getElementById('search_criteria').selected_car_types);update_car_type_list();return false;">
                      <img border="0" src="images/down_button.GIF" width="24" height="24"  alt="Click this down arrow to modify the display order of the car types" ></a></font></td>
                  <td></td> 
                    <td align="right">
                    
                    <select name="unselected_companies" size="12" style="width:175px" multiple>
                      <% Dim Companies                                       %> 
                      <% Dim SelectedCompanies							     %>
                      <% Companies = ""                                      %>
                      <% SelectedCompanies = Companies						 %> 
                      <% While adoRS2.EOF = False                            %>
                       	<option value="<%=adoRS2.Fields("vendor_cd").Value %>"><%=adoRS2.Fields("vendor_name_rpt").Value %> (<%=adoRS2.Fields("vendor_cd").Value%>)</option>
 					   <%	   adoRS2.MoveNext 
					   Wend 
					  %>
					</select>
					
					</td>
                    <td>
                      
                    <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').unselected_companies, document.getElementById('search_criteria').selected_companies, false );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_right.GIF" width="24" height="22"  alt="Add the highlighted companies types"   ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').unselected_companies, document.getElementById('search_criteria').selected_companies, true );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_right_all.GIF" width="24" height="22"  alt="Add all the companies"   ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').selected_companies, document.getElementById('search_criteria').unselected_companies, false );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_left.GIF" width="24" height="22"  alt="Remove the highlighted companies"   ></a><br>
                      
    
                    <a href="javascript:void(0)" onclick="moveDualList( document.getElementById('search_criteria').selected_companies, document.getElementById('search_criteria').unselected_companies, true );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_left_all.GIF" width="24" height="22"  alt="Remove all the companies"   ></a></td>
                    <td>
                      
                      <select name="selected_companies" id="selected_companies" size="12" style="width:175px" multiple>
                      <% Dim CompanyArray                                            %> 
                      <% Dim CompanyCodeList                                         %> 
                      <% CompanyArray = Split(SelectedCompanies, ",")                %>
                      <% For intIndex = 0 To UBound(CompanyArray) - 1 Step 2         %>
                        <option value="<%=CompanyArray(intIndex) %>"><%=CompanyArray(intIndex + 1) %></option>
                        <% CompanyCodeList = CompanyCodeList & CompanyArray(intIndex) & ","  %>                                        
                      <% Next %>
                      </select></td>
                    <td>
                      <a href="javascript:void(0)" onclick="moveOptionUp(document.getElementById('search_criteria').selected_companies);update_company_list(true, '');return false;">
                      <img border="0" src="images/up_button.GIF" width="24" height="24" alt="Click this up arrow to modify the order" ></a><br>
                    <font color="#080000">
                    <a href="javascript:void(0)" onclick="moveOptionDown(document.getElementById('search_criteria').selected_companies);update_company_list(true, '');return false;">
                      <img border="0" src="images/down_button.GIF" width="24" height="24"  alt="Click this down arrow to modify the order" ></a></font></td>
                    <td></td>
                  <td valign="top">              
                <select name="highlighted_company" id="highlighted_company" style="width:200px" size="12">
                <% strHighlightVendor = ""                                       %>
	                <option value=""></option>
                </select></td>
                      <td width="30"></td>
                 </tr>


                 <tr>
                     <td colspan="9" align="left"><br /><input type="submit" name="submitform" id="submitform" value="Update Selected Profiles" /></td>
                 </tr>
                </table>
                </td>
                </tr>

            </table>

 <table align="center" border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4" id="table3">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
 </table>

    <table align="center" width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td style="width: 80%" ><font size="2"> |
    <a href="javascript:checkAll(document.search_criteria.profile_id)"  title="Select and check all the profiles on this page" >Select All</a> | 
    <a  href="javascript:uncheckAll(document.search_criteria.profile_id)" title="Unselect and un-check all the profiles on this page" >Unselect All</a> |
    <a href="search_profiles_car.asp">Search Profiles</a> | 
    <a href="#bottom">Go to bottom</a></font> |</td>
  </tr>
      </table>
	  
 <table align="center" border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4" id="table3">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
 </table>
 
             <table align="center" border="1" width="1110: cellpadding="2" style="border-collapse: collapse"  bordercolor="#FFFFFF" background="images/alt_color.gif">
            <tr>
                <td class="profile_header" style="background-color: #E07D1A"><b>Selected</b></td>
                <td class="profile_header"><b>Profile Description</b></td>
                <td class="profile_header"><b>Car Type Codes</b></td>
                <td class="profile_header"><b>Car Companies</b></td>
            </tr>
<%
    

    Rem Display the profiles 
    While adoRS.EOF = False  
   		If strClass <> "profile_dark" Then
        	strClass = "profile_dark" 'background-color:#B2BEC4
        Else
        	strClass = "profile_light" 'background-color:#CFD7DB
        End If
        Response.Write "<tr>" & vbcrlf
        Response.Write "<td align='center' bgcolor='#FDC677'><input type='checkbox' name='profile_id' value='" & adoRS.Fields("profile_id") & "'></td>" & vbcrlf
        Response.Write "<td class='" & strClass & "' ><a href='search_criteria_car.asp?profile=" & adoRs.Fields("profile_id").Value & "'>" & adoRS.Fields("desc").Value & "</a></td>" & vbcrlf
        Response.Write "<td class='" & strClass & "' >" & adoRS.Fields("shop_car_type_cds").Value & "</td>" & vrcrlf
        Response.Write "<td class='" & strClass & "' >" & adoRS.Fields("vend_cds").Value & "</td>" & vbcrlf
        Response.Write "</tr>" & vbcrlf
    
    
        adoRS.MoveNext
    Wend
    
    
 %>

        </table>
 </form>
 
     <table align="center" border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4" id="table4">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
 </table>
       
    <table align="center" width="1108" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td style="width: 80%" ><font size="2"> |
    <a name="bottom" href="javascript:checkAll(document.search_criteria.profile_id)"  title="Select and check all the profiles on this page" >Select All</a> | 
    <a  href="javascript:uncheckAll(document.search_criteria.profile_id)" title="Unselect and un-check all the profiles on this page" >Unselect All</a> |
    <a href="search_profiles_car.asp">Search Profiles</a> | 
    <a href="#top">Go to Top</a></font> |</td>
  </tr>
      </table>

     <table align="center" border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4" id="table5">
    <tr>
      <td background="images/ruler.gif"></td>
    </tr>
 </table>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<%	
	Rem Clean up
	Set adoCmd  = Nothing 
	Set adoCmd2 = Nothing 
	Set adoCmd5 = Nothing 	

	Set adoRS  = Nothing 
	Set adoRS2 = Nothing 
	Set adoRS5 = Nothing 
%>