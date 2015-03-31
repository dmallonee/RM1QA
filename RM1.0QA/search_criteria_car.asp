<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/login_check_ex.asp" -->
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
	Dim blnAllowSaving

	Dim adoPrices
	Dim strUserId
	Dim strHighlightVendor
	
	On Error Resume Next
	
	strUserId     = Request.Cookies("rate-monitor.com")("user_id")
	strSelfVendCd = Request.Cookies("rate-monitor.com")("vend_cd")
	intRptLimit   = Request.Cookies("rate-monitor.com")("rpt_limit")
		
	Session("user_id") = strUserID
	strConn = Session("pro_con")
    intProfileStatus = Request("profile_status")
	
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_shop_profile_select"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Null)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 1024)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@enabled", 3, 1, 0, intProfileStatus)
		
	Set adoRS = adoCmd.Execute
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting profile information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If

		
	Rem Get the data sources
	Set adoCmd1 = CreateObject("ADODB.Command")

	adoCmd1.ActiveConnection =  strConn
	adoCmd1.CommandText = "data_source_select"
	adoCmd1.CommandType = 4

	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@lob_id",  3, 1, 0, 2)
	adoCmd1.Parameters.Append adoCmd1.CreateParameter("@user_id", 3, 1, 0, strUserId)

	Set adoRS1 = adoCmd1.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting data source information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If


	Rem Get the vendors
	Set adoCmd2 = CreateObject("ADODB.Command")

	adoCmd2.ActiveConnection =  strConn
	adoCmd2.CommandText = "vendor_select"
	adoCmd2.CommandType = 4
	
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_cd",   200, 1,  2, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@vendor_name", 200, 1, 50, Null)
	adoCmd2.Parameters.Append adoCmd2.CreateParameter("@user_id",       3, 1,  0, strUserId)
		
	Set adoRS2 = adoCmd2.Execute


	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting vendor information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If

	
	Rem Get the vendors
	Set adoCmd3 = CreateObject("ADODB.Command")

	adoCmd3.ActiveConnection =  strConn
	adoCmd3.CommandText = "user_city_select"
	adoCmd3.CommandType = 4
	
	adoCmd3.Parameters.Append adoCmd3.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS3 = adoCmd3.Execute
	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting city information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
	
	
	Rem Get the car types
	Set adoCmd5 = CreateObject("ADODB.Command")

	adoCmd5.ActiveConnection =  strConn
	adoCmd5.CommandText = "car_type_select"
	adoCmd5.CommandType = 4
	
	adoCmd5.Parameters.Append adoCmd5.CreateParameter("@user_id", 3, 1, 0, strUserId)
		
	Set adoRS5 = adoCmd5.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while collecting car type information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
	
	Rem Initialize it
	ProfileID = 0
	
	If (Request("profile") > 0) Then
	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_profile_select"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Append adoCmd.CreateParameter("@desc", 200, 1, 255)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds", 200, 1, 1024)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id", 3, 1, 0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Request("profile"))
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 1024)
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@enabled", 3, 1, 0, intProfileStatus)

		Set adoRS4 = adoCmd.Execute
		
		ProfileID = adoRS4.Fields("profile_id").Value
		
		blnProfileLoad = True

		If err.number <> 0 Then
		   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
		   response.write "<b>An error cccured while collecting specific profile information</b><br>"
		   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
		   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
		   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
		   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
		   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

		End If
		

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


		If err.number <> 0 Then
REM		   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
REM		   response.write "<b>An error cccured while collecting default profile information</b><br>"
REM		   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
REM		   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
REM		   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
REM		   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
REM		   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
		End If

		
	End If

	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "schedule_group_select"
	adoCmd.CommandType = 4

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",    3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id", 3, 1, 0, Request("profile"))

	Set adoRS6 = adoCmd.Execute


	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "currency_code_select"
	adoCmd.CommandType = 4

	Set adoRSCurrency = adoCmd.Execute


	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "country_select"
	adoCmd.CommandType = 4

	Set adoRSCountry = adoCmd.Execute




	If err.number <> 0 Then
REM	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
REM	   response.write "<b>An error cccured</b><br>"
REM	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
REM	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
REM	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
REM	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
REM	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	End If
	
	If (adoRS4.Fields("user_id").Value = CInt(strUserId)) Or (adoRS4.Fields("user_id").Value = 0) Then
		blnAllowSaving = True
	Else
		blnAllowSaving = False
	End If

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<META http-equiv="Content-Language" content="en-us">
<META HTTP-EQUIV="refresh" CONTENT="900;URL=default_session.asp">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Criteria</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script type="text/javascript" language="JavaScript" src="inc/ts_picker.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/center.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/date_check.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/multiple_select_support.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/multiple_select_support2.js"></script>
<script type="text/javascript" language="JavaScript" src="inc/selectbox.js"></script>
<script type="text/javascript" language="Javascript">


    function not_enabled() {
        alert("This section is currently unavailable." + "\n" + "Please contact your Rate-Highway rep if you would like to enable it");
        //return true;
    }


    function update_company_list(reset_highlight, discount_list) {

        var list = new String("")
        var discount = discount_list.split('|')
        // the profile save saves them in reverse order, so we need to flip them back into order
        discount.reverse();


        if (reset_highlight == true) {
            removeAllOptions(document.search_criteria.highlighted_company);
            addOption(document.search_criteria.highlighted_company, "(please select one)", "", true);
        }
        // Clear out the discount codes
        deleteRows();

        for (i = 0; i < document.search_criteria.selected_companies.options.length; i++) {
            if (reset_highlight == true) {
                addOption(document.search_criteria.highlighted_company, document.search_criteria.selected_companies.options[i].text, document.search_criteria.selected_companies.options[i].value, false);
            }

            if (discount_list == '') {
                addRowToTable(document.search_criteria.selected_companies.options[i].text, '');
            }
            else {
                addRowToTable(document.search_criteria.selected_companies.options[i].text, discount[((i * 2) + 1)]);
            }

            if (i == 0) {
                list = document.search_criteria.selected_companies.options[i].value;
            }
            else {
                list = list + ',' + document.search_criteria.selected_companies.options[i].value;
            }
        }

        document.search_criteria.company_list.value = list;
        //document.search_criteria.selected_companies.options.text; 

    }
    function update_car_type_list() {

        var list = new String("")

        for (i = 0; i < document.search_criteria.selected_car_types.options.length; i++) {
            if (i == 0) {
                list = document.search_criteria.selected_car_types.options[i].value
            }
            else {
                list = list + ',' + document.search_criteria.selected_car_types.options[i].value
            }
        }

        document.search_criteria.car_type_list.value = list
        //document.search_criteria.selected_companies.options.text; 

    }

    function update_city_list() {

        var list = new String("")

        for (i = 0; i < document.search_criteria.selected_cities.options.length; i++) {
            if (i == 0) {
                list = document.search_criteria.selected_cities.options[i].value
            }
            else {
                list = list + ',' + document.search_criteria.selected_cities.options[i].value
            }
        }


        document.search_criteria.cities_list.value = list
        return true
    }


    function check_all_days(fieldName) {

        thisButton = document.forms[0][fieldName];
        for (var i = 0; i < thisButton.length; i++) {
            thisButton[i].checked = true;
        }

    }

    function deleteRows() {
        // Check all rows in table. Length of table will change during loop if deletes occur.
        var tbl = document.getElementById('discount_codes');
        for (i = tbl.rows.length - 1; i != 0; i--) {
            tbl.deleteRow(i);
        } // End process all rows
    } // End deleteRows function
</script>
<script language="JavaScript" type="text/JavaScript">

    function CopyCityCode() {
        //document.search_criteria.return_city.value = document.search_criteria.pickup_city.options[document.search_criteria.pickup_city.selectedindex].value;
        selected = document.search_criteria.pickup_city.selectedIndex;
        fieldValue = document.search_criteria.pickup_city.options[selected].value;
        document.search_criteria.return_city.value = fieldValue;

        return;
    }

    function disp_text() {
        var w = document.search_criteria.profile.selectedIndex;
        var selected_text = document.search_criteria.profile.options[w].text;
        document.search_criteria.profile_text.value = selected_text;
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

        if (SubmitType == 'open') {

            //verify that a profile has been selected
            selected = document.search_criteria.profile.selectedIndex;
            fieldValue = document.search_criteria.profile.options[selected].value;
            fieldText = document.search_criteria.profile.options[selected].text;

            if (fieldValue == 0) {
                alert("You must select a valid profile before you can open it");
                return false;

            }
            else {
                document.search_criteria.action = 'search_criteria_car.asp';
                //document.search_criteria.submit;
                return true;
            }
        }

        else if (SubmitType == 'save') {

            // This needs to be cleared out so that it is not confused with a save as
            document.search_criteria.profile_save_as.value = ""

            //verify that a profile has been selected
            selected = document.search_criteria.profile.selectedIndex;
            fieldValue = document.search_criteria.profile.options[selected].value;
            fieldText = document.search_criteria.profile.options[selected].text;

            if (fieldValue == 0) {
                alert("You must select and open a profile before you can save it");
                return false;

            }


            else {
                if (confirm("Are you sure you want to overwrite " + fieldText + "?")) {
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

        else if (SubmitType == 'saveas') {

            //verify that a profile has been named
            validChars = "abcdefghijklmnopqrstuvwxyz";
            validChars += "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            validChars += "0123456789";

            fieldName = document.search_criteria.profile_save_as;
            fieldValue = fieldName.value;
            fieldLength = fieldValue.length;
            minLength = 1;
            maxLength = 255;

            var err01 = "Non-valid character(s) found in the profile save as name, please no special characters.";
            var err03 = "Please enter a profile name with at least " + minLength + " character in length.";
            var err04 = "Please enter less than " + maxLength + " characters in length for the profile name.";

            if (fieldValue == "") {
                alert(err03);
                fieldName.focus();
                return false;
            }
            else if (fieldLength < minLength) {
                alert(err03);
                fieldName.focus();
                return false;
            }
            else if ((fieldLength > maxLength) && (maxLength > 0)) {
                alert(err04);
                fieldName.focus();
                return false;
            }

            else {
                for (var i = 0; i < fieldLength; i++) {
                    if (validChars.indexOf(fieldValue.charAt(i)) == -1) {
                        alert(err01);
                        fieldName.focus();
                        return false;
                    }
                    else {
                        if (ValidateForm() == true) {
                            document.search_criteria.action = 'search_profile_insert_car.asp';
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

        else if (SubmitType == 'searchnow') {

            if (confirm("Are you sure you want to search?")) {
                if (ValidateForm() == true) {
                    document.search_criteria.action = 'search_request_insert_car.asp';
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

        doSubmit = new Boolean()
        doSubmit = true

        if (doSubmit == true) {
            var prevdate = document.getElementById("rental_begin_date").value;
            var thisdate = document.getElementById("rental_end_date").value;
            var date0 = new Date(prevdate);
            var date1 = new Date(thisdate);
            var datediff2 = Math.ceil(date0.getTime() - date1.getTime())
            if (datediff2 > 0) {
                document.getElementById("rental_end_date").value = prevdate;
                alert("Last date cannot be earlier than first date");
                document.getElementById("rental_end_date").focus();
                doSubmit = false;
            }
        }

        if (doSubmit == true) {
            if (document.search_criteria.selected_cities.options.length < 1) {
                alert("Please select at least one city code to search.");
                document.search_criteria.selected_cities.focus();
                doSubmit = false;
            }
        }


        if (doSubmit == true) {
            //alert("entered ValidateForm");
            //DOW list
            selection = null;
            thisButton = document.search_criteria.dow_list;
            for (var i = 0; i < thisButton.length; i++) {
                if (thisButton[i].checked) {
                    selection = thisButton[i].value;
                }
            }

            if (selection == null) {
                alert("Please check at least one of the day of week boxes.");
                doSubmit = false

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

        if (doSubmit == true) {
            if (document.search_criteria.begin_date.value == "") {
                alert("Please enter a valid begin date");
                document.search_criteria.begin_date.focus();
                doSubmit = false;
            }
        }


        if (doSubmit == true) {
            if (document.search_criteria.end_date.value == "") {
                alert("Please enter a valid end date");
                document.search_criteria.end_date.focus();
                doSubmit = false;
            }
        }

        if (doSubmit == true) {
            if (document.search_criteria.selected_car_types.options.length == 0) {
                alert("Please select at least one car type");
                document.search_criteria.unselected_car_types.focus();
                doSubmit = false;
            }
        }

        if (doSubmit == true) {
            if (document.search_criteria.selected_companies.options.length == 0) {
                alert("Please select at least one car company");
                document.search_criteria.unselected_companies.focus();
                doSubmit = false;
            }
        }

        if (doSubmit == true) {
            selected = document.search_criteria.highlighted_company.selectedIndex;
            fieldValue = document.search_criteria.highlighted_company.options[selected].value;

            if (fieldValue == "") {
                alert("Please select a company for highlight");
                document.search_criteria.highlighted_company.focus();
                doSubmit = false;
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



    function showCodeOption(obj, site) {

        if (typeof (obj) == "object") {
            site = obj.options[obj.selectedIndex].value;
        }

        // Travelocity Deep, Expedia Deep or Special and Avis all need to be able to input discount codes.
        if (site == 'TRD') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'EXS') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'EXD') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'VZI') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'M06') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'M07') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'M08') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'ECS') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else if (site == 'WSA') {
            document.getElementById('discount_area').style.display = 'block';
        }
        else {
            document.getElementById('discount_area').style.display = 'none';
        }
    }

    function addRowToTable(company_name, discount_code) {
        var tbl = document.getElementById('discount_codes');
        var lastRow = tbl.rows.length;
        // if there's no header row in the table, then iteration = lastRow + 1 otherwise it is just lastRow
        var iteration = lastRow;
        var row = tbl.insertRow(lastRow);
        var y = row.insertCell(0)
        var z = row.insertCell(1)
        var v1 = row.insertCell(2)
        var v2 = row.insertCell(3)

        var el = document.createElement('input');
        el.type = 'text';
        el.name = 'discount_code';
        el.id = 'discount_code';
        el.value = discount_code;
        el.size = 20;
        v1.appendChild(el);

        var el = document.createElement('input');
        el.type = 'text';
        el.name = 'discount_code_vendor';
        el.id = 'discount_code_vendor';
        el.size = 20;
        el.value = company_name;
        el.readOnly = true;
        el.style.backgroundColor = "#D8DEE1";
        v2.appendChild(el);
    }

</script>



<script language='Javascript' type="text/javascript" >
    function centerPopUp(url, name, width, height, scrollbars) {

        if (scrollbars == null) scrollbars = "0"

        str = "";
        str += "resizable=1,";
        str += "scrollbars=" + scrollbars + ",";
        str += "width=" + width + ",";
        str += "height=" + height + ",";

        if (window.screen) {
            var ah = screen.availHeight - 30;
            var aw = screen.availWidth - 10;

            var xc = (aw - width) / 2;
            var yc = (ah - height) / 2;

            str += ",left=" + xc + ",screenX=" + xc;
            str += ",top=" + yc + ",screenY=" + yc;
        }
        window.open(url, name, str);
    } 

</script> 


<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all">
<script language="javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>	
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

<body onload="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_rules_on.gif','images/b_user_on.gif','images/b_system_on.gif');">


<!--
document.getElementById('discount_area').style.display = 'none';
-->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif">
    <img src="images/top.jpg" width="770" height="91" alt=""></td>
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
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
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
      <td colspan="1">&nbsp;</td>
    </tr>
    <tr height="1">
      <td bgcolor="#000000" colspan="1" height="1"><img src="pixel.gif" width="1" height="1"></td>
    </tr>
  </table>
    <table align="center" cellpadding="0" cellspacing="0" border="0" width="769" bgcolor="#CCCCFF">
      <tr>
        <td width="1" bgcolor="#000000">
        <img src="pixel.gif" width="1" height="1"></td>
        <td bgcolor="#D9DEE1" width="768">
        <table border="0" cellspacing="5" cellpadding="5" width="768">
          <tr>
            <td bgcolor="#FFFFFF" width="748"><font color="#080000">
            <!-- JUSTTABS TOP CLOSE -->
            <table width="745" border="0" cellspacing="0" cellpadding="2">
              <tr valign="bottom">
                <td width="162" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="93" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="90" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="52" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="83" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="144" bgcolor="#FFFFFF">&nbsp;</td>
                <td bgcolor="#FFFFFF" width="93" align="left" valign="middle" >Help 
                <img border="0" alt="Click to view Help" src="images/help_button.jpg" width="32" height="32" onclick="centerPopUp('help_search_criteria.htm', 'help', 650, 400, 1)"></td>
              </tr>
              <tr valign="bottom">
                <td width="162" bgcolor="#FFFFFF">
                <img src="images/ti_profile.gif" width="162" height="25"></td>
                <td width="93" bgcolor="#FFFFFF">
                <button name="open" style="height: 25px; width: 90px" value="open" onclick="return confirmSubmit('open')" type="submit">
                Open</button></td>
                <td width="90" bgcolor="#FFFFFF">
                <% If blnAllowSaving Then %>
                <button name="save" style="height: 25; width: 90" value="save" onclick="return confirmSubmit('save')" type="submit">Save</button>
                <% Else %>
                <button name="save" style="height: 25; width: 90" value="save" onclick="return confirmSubmit('save')" type="submit" disabled="disabled">Save</button>
				<% End If %>
                </td>
                <td width="52" bgcolor="#FFFFFF"></td>
                <td width="83" bgcolor="#FFFFFF"></td>
                <td width="144" bgcolor="#FFFFFF">
                <% If blnAllowSaving Then %>
                <button name="save_as" style="height: 25; width: 90" value="save_as" type="submit" onclick="return confirmSubmit('saveas')">Save As</button>
                <% Else %>
                <button name="save_as" style="height: 25; width: 90" value="save_as" type="submit" onclick="return confirmSubmit('saveas')" disabled="disabled">Save As</button>
				<% End If %>
                </td>
                <td bgcolor="#FFFFFF" width="93">&nbsp;</td>
              </tr>
              <tr valign="bottom">
                <td width="162" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="150" colspan="3" bgcolor="#FFFFFF">
                <% Dim strProfileText         %>
                <% strProfileText = "Default" %>
                <select name="profile" style="width:230;" >
                <option value="0">Default</option>
                <% While adoRS.EOF = False %> 
	                <% If ProfileID = adoRS.Fields("profile_id").Value Then %>
	                	<option selected value="<%=adoRS.Fields("profile_id").Value %>"><%=adoRS.Fields("desc").Value %></option>
	                	<% strProfileText = adoRS.Fields("desc").Value %>
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
                <input name=profile_text type="hidden" value="<%=strProfileText          %>" >
				</td>
                <td width="83" bgcolor="#FFFFFF">&nbsp;</td>
                <td width="316" colspan="2" bgcolor="#FFFFFF">
                <% If blnAllowSaving Then %>
                <input type="text" name="profile_save_as" size="36" onfocus="this.className='focus';cl(this,'save profile as...');" onblur="this.className='';fl(this,'save profile as...');" >
                <% Else %>
                <input type="text" name="profile_save_as" size="36" onfocus="this.className='focus';cl(this,'save profile as...');" onblur="this.className='';fl(this,'save profile as...');" disabled="disabled" >
				<% End If %>
                
                </td>
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
                
                <img src="images/ti_city_location.gif" width="162" height="25"></td>
                <td width="24" valign="top">
                <img src="images/separator.gif" width="20" height="20"></td>
                <td width="547" valign="bottom">
                <table width="475" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table1">
                  <tr>
                    <td width="475" colspan="3">
                    <div align="center">
                        (use buttons to move between selected/unselected)
                        <br>
                        Unselected cities&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                        Selected cities&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 
                    </div>
                    </td>
                  </tr>
                  <tr>
                    <td width="209">
                    
                    <select name="unselected_cities" size="5" style="width:200;" multiple >
                      <% Dim Cities                                          %> 
                      <% Dim SelectedCities  							     %>
                      <% Cities = adoRS4.Fields("city_cd").Value & ","    		 %> 
                      <% While adoRS3.EOF = False                            %>
                      <% 	If (InStr(Cities, adoRS3.Fields("city_cd").Value & ",") = 0) Or (Cities = "") Then %>
                      			<option value="<%=adoRS3.Fields("city_cd").Value %>"><%=adoRS3.Fields("city_cd").Value %></option>
                      <% 	Else 											 %>		                    
		              <%     	SelectedCities = SelectedCities & TRIM(adoRS3.Fields("city_cd").Value) & ","  %>
						
					  <% 	End If 											 %>
					  <%    adoRS3.MoveNext 								 %>
					  <% Wend												 %> 
					  <% Set adoRS3 = Nothing 								 %>
					</select></td>
                    <td width="37">
                      <img border="0" src="images/move_right.GIF" width="24" height="22"  onclick="moveDualList( document.search_criteria.unselected_cities, document.search_criteria.selected_cities, false );update_city_list();return false" ><br>
                      <img border="0" src="images/move_right_all.GIF" width="24" height="22"  onclick="moveDualList( document.search_criteria.unselected_cities, document.search_criteria.selected_cities, true );update_city_list();return false"  ><br>
                      <img border="0" src="images/move_left.GIF" width="24" height="22"  onclick="moveDualList( document.search_criteria.selected_cities, document.search_criteria.unselected_cities, false );update_city_list();return false"  ><br>
                      <img border="0" src="images/move_left_all.GIF" width="24" height="22"  onclick="moveDualList( document.search_criteria.selected_cities, document.search_criteria.unselected_cities, true );update_city_list();return false"  >
                    </td>
                    <td width="228">
                      
                      <select name="selected_cities" size="5" style="width:200;" multiple>
                      <% Dim CityArray                                            %> 
                      <% CityArray = Split(SelectedCities, ",")                   %>
                      <% For intIndex = 0 To UBound(CityArray) - 1                %>
                        <option value="<%=CityArray(intIndex) %>"><%=CityArray(intIndex) %></option>
                      <% Next %>
                      </select></td>
                  </tr>
                </table>
                </td>
              </tr>
              <tr valign="bottom">
                <td width="163" bgcolor="#FFFFFF">
                &nbsp;</td>
                <td width="24" bgcolor="#FFFFFF">
                <p align="left">
                &nbsp;</td>
                <td width="547" bgcolor="#FFFFFF">
                <font color="#080000">
                Return city (one-way searches only):
				<% If IsNull(adoRS4.Fields("rtrn_city_cd").Value) Then %>
                <input type="text" name="return_city" size="6" onfocus="this.className='focus';cl(this,'(same)');" onblur="this.className='';fl(this,'(same)');" value="(same)">
				<% Else %>
                <input type="text" name="return_city" size="6" onfocus="this.className='focus';cl(this,'(same)');" onblur="this.className='';fl(this,'(same)');" value="<%=adoRS4.Fields("rtrn_city_cd").Value %>">
                <% End If %>
                
                <% If IsNull(adoRS4.Fields("oneway_reverse").Value) Then %>
                <input name="oneway_reverse" type="checkbox" id="oneway_reverse"  value="True"><label for="oneway_reverse"> One-way reverse</label>
				<% ElseIf CBool(adoRS4.Fields("oneway_reverse").Value) Then %>
                <input name="oneway_reverse" type="checkbox" id="oneway_reverse"  value="True" checked="checked" ><label for="oneway_reverse"> One-way reverse</label> 
				<% Else   %>
                <input name="oneway_reverse" type="checkbox" id="oneway_reverse"  value="True"><label for="oneway_reverse"> One-way reverse</label> 
                <% End If %>
                </font>
                </td>


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
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <img src="images/separator.gif" width="745" height="15"></td>
              </tr>
            </table>
            <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
			  <% If intRptLimit > 0 Then %>
              <tr valign="bottom">
                <td width="162">
                &nbsp;</td>
                <td width="132">
                &nbsp;</td>
                <td width="88">
                &nbsp;<td width="22" valign="middle" >
				&nbsp;</td>
                <td width="21">
                &nbsp;</td>
                <td width="226" colspan="3">
                <p align="center"><font size="1" color="#080000">Your report may not exceed <%=intRptLimit %> days.<br>
                Requests that do will be auto-adjusted.</font></td>
                <td width="54">&nbsp;</td>
              </tr>
			  <% End If %>
              <tr valign="bottom">
                <td width="162">
                <img src="images/rental%20dates.gif" width="162" height="25"></td>
                <td width="132">
                Begin: 
                (mm/dd/yy)</td>
                <td width="88">
                <!--
                <input type="text" name="begin_date" class="fsmall" style="width:85" size="10" maxlength="10" onfocus="javascript:vDateType='1'" onkeyup="DateFormat(this,this.value,event,false,'1')" onblur="DateFormat(this,this.value,event,true,'1')">
                -->
				<% Dim datBeginDate_value
				   Dim datEndDate_value
				%>
				
                <% If adoRS4.Fields("exact_dates").Value = 0 Then %> 
                
                	<% datBeginDate_value = FormatDateTime(DateAdd("d", Now, DateDiff("d", adoRS4.Fields("change_dttm").Value, adoRS4.Fields("begin_arv_dt").Value)), 2) %>

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
	                <% datBeginDate_value = adoRS4.Fields("begin_arv_dt").Value %>
	 	
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
                </td>
                <td width="21">
                <img src="images/separator.gif" width="20" height="27"> </td>
                <td width="124">
                End: 
                (mm/dd/yy)</td>
                <td width="82">
                <!-- 
              <input type="text" name="end_date" class="fsmall" style="width:85" size="10" maxlength="10" onfocus="javascript:vDateType='1'" onkeyup="DateFormat(this,this.value,event,false,'1')" onblur="DateFormat(this,this.value,event,true,'1')">
              -->
                <% If adoRS4.Fields("exact_dates").Value = 0 Then %> 
	                <% datEndDate_value = FormatDateTime(DateAdd("d", Now, DateDiff("d", adoRS4.Fields("change_dttm").Value, adoRS4.Fields("end_arv_dt").Value)), 2) %>

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
                		<% datEndDate_value = adoRS4.Fields("end_arv_dt").Value %>
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
                ABP(Rolling) 
                or Fixed: </td>
                <td width="450">
                <select name="exact_dates" style="width:230" size="1">
                <% If adoRS4.Fields("exact_dates").Value = False Then 	%>
	                <option value="1">Fixed Dates (those exact dates)</option>
	                <option selected value="0">Rolling Dates (# days from today)</option>
                <% ElseIf adoRS4.Fields("exact_dates").Value = True Then %>
	                <option selected value="1">Fixed Dates (those exact dates)</option>
	                <option value="0">Rolling Dates (# days from today)</option>
                <% Else %>
	                <option value="1">Fixed Dates (those exact dates)</option>
	                <option selected value="0">Rolling Dates (# days from today)</option>
                <% End If %>

            

                </select></td>
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
                
                <% If Instr(adoRS4.Fields("dow_list").Value, "1") Then %>
                	<input type="checkbox" name="dow_list" value="1" checked> 
                <% Else %>
                	<input type="checkbox" name="dow_list" value="1"> 
                <% End If %>
                </td>
                <td width="27">
                Sun
                </td>
                <td width="25">
                <% If Instr(adoRS4.Fields("dow_list").Value, "2") Then %>
                <input type="checkbox" name="dow_list" value="2" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="2"> <% End If %>
                </td>
                <td width="30">
                Mon
                </td>
                <td width="22">
                <% If Instr(adoRS4.Fields("dow_list").Value, "3") Then %>
                <input type="checkbox" name="dow_list" value="3" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="3"> <% End If %>
                </td>
                <td width="26">
                Tue
                </td>
                <td width="24">
                <% If Instr(adoRS4.Fields("dow_list").Value, "4") Then %>
                <input type="checkbox" name="dow_list" value="4" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="4"> <% End If %>
                </td>
                <td width="31">
                Wed
                </td>
                <td width="24">
                <% If Instr(adoRS4.Fields("dow_list").Value, "5") Then %>
                <input type="checkbox" name="dow_list" value="5" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="5"> <% End If %>
                </td>
                <td width="24">
                Thu
                </td>
                <td width="22">
                <% If Instr(adoRS4.Fields("dow_list").Value, "6") Then %>
                <input type="checkbox" name="dow_list" value="6" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="6"> <% End If %>
                </td>
                <td width="17">
                Fri</td>
                <td width="210">
                <% If Instr(adoRS4.Fields("dow_list").Value, "7") Then %>
                <input type="checkbox" name="dow_list" value="7" checked> <% Else %>
                <input type="checkbox" name="dow_list" value="7"> <% End If %>
                Sat</td>
              </tr>
            </table>

<%
    'if(!string.IsNullOrEmpty(adoRS4["shop_boundary_start"].ToString())) then
        boundaryStartDate_value = adoRS4.Fields("shop_boundary_start").Value
    'end
    'if (!string.IsNullOrEmpty(adoRS4["shop_boundary_end"].ToString()) then
        boundaryEndDate_value = adoRS4.Fields("shop_boundary_end").Value
    'end if
    if (Request.Cookies("rate-monitor.com")("shop_boundary") = "True") then
 %>
            <table width="745" border="0" cellspacing="0" cellpadding="0" background="images/alt_color.gif">
                <tr>
                <td width="179" align="left"></td>
                <td align="left"><br />
                    Boundary Start Date:
                        <input type="text" name="boundary_start_date" id="boundary_start_date" class="cb_txtdate" value="<%=boundaryStartDate_value%>" onfocus="openCal(this,'boundary_start_date','boundary_end_date','calbox','torowed','us','vertical');" onclick="clearTimeout(t_calcloser)" style="width:80px" />
                    Boundary End Date:
                        <input type="text" name="boundary_end_date" id="boundary_end_date" class="cb_txtdate" value="<%=boundaryEndDate_value%>" onfocus="openCal(this,'boundary_start_date','boundary_end_date','calbox','torowed','us','vertical');" onclick="clearTimeout(t_calcloser)" style="width:80px" /><br />
                        Dates will not be shopped before the boundary start and after the boundary end.<br>
			If there is no end date, leave that field blank.
                </td>



                </tr>
            </table>
<% else 'we have to retain the values%>
                        <input type="hidden" name="boundary_start_date" id="Text1" class="cb_txtdate" value="<%=boundaryStartDate_value%>" />
                        <input type="hidden" name="boundary_end_date" id="Text2" class="cb_txtdate" value="<%=boundaryEndDate_value%>" />
<%end if %>



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
                
                <img src="images/ti_details.gif" width="162" height="25"></td>
                <td width="97">
                &nbsp;Pick-up 
                Time :</td>
                <td width="92">
                <select name="arrival_time" style="width:90" size="1"> 
                <% Dim strTime
                   Dim strTimeValue 
                   Dim strCheckTime
                   Dim strTimeSelected
                   
                   strCheckTime = trim(adoRS4.Fields("arv_tm").Value)
                   

                %>
                	 <!--
                	 <option value='00:00'>Midnight</option>
                	 -->
                <%

                   For intIndex = 0 To 11  
                     strTime = intIndex & ":00 am"
                     strTimeValue = intIndex & ":00"
                     
                     If intIndex = 0 Then
                     	strTime = "Midnight"
                     End If
                     
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":15 am"
                     
                     If intIndex = 0 Then
                     	strTime = "12:15 am"
					 End If	                 
                     
                     strTimeValue = intIndex & ":15"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%

                     strTime = intIndex & ":30 am"

                     If intIndex = 0 Then
                     	strTime = "12:30 am"
					 End If	                 
                     
                     strTimeValue = intIndex & ":30"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%

                     strTime = intIndex & ":45 am"
                     
                     If intIndex = 0 Then
                     	strTime = "12:45 am"
					 End If	                 
                     
                     strTimeValue = intIndex & ":45"
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
                <!--
                	 <option selected value='12:00'>Noon</option>
                -->
                <%
				   Else
                %>
                <!--
                	 <option value='12:00'>Noon</option>
                -->
                <%
				   End If

                   For intIndex = 0 To 11  
                     strTime = intIndex & ":00 pm"
                     strTimeValue = intIndex + 12 & ":00"
                     
                     If strTime = "0:00 pm" Then
                     	strTime = "Noon"
					 End If	                 
                     
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":15 pm"

                     If intIndex = 0 Then
                     	strTime = "12:15 pm"
					 End If	                 

                     strTimeValue = intIndex + 12 & ":15"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":30 pm"

                     If intIndex = 0 Then
                     	strTime = "12:30 pm"
					 End If	                 
                     
                     strTimeValue = intIndex + 12 & ":30"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":45 pm"

                     If intIndex = 0 Then
                     	strTime = "12:45 pm"
					 End If	                 
                     
                     strTimeValue = intIndex + 12 & ":45"
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
                <td width="99">Drop-off 
                Time : </td>
                <td width="130">
                &nbsp;<select name="return_time" style="width:90">
		        <% strCheckTime = trim(adoRS4.Fields("rtrn_tm").Value)

                %>
                <!--
                	 <option value='00:00'>Midnight</option>
                -->
                <%

                   For intIndex = 0 To 11  
                     strTime = intIndex & ":00 am"
                     
                     If intIndex = 0 Then
                     	strTime = "Midnight"
					 End If	                 

                     
                     strTimeValue = intIndex & ":00"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":15 am"
                     
                     If intIndex = 0 Then
                     	strTime = "12:15 am"
					 End If	                 
                     
                     strTimeValue = intIndex & ":15"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":30 am"

                     If intIndex = 0 Then
                     	strTime = "12:30 am"
					 End If	                 
                     
                     strTimeValue = intIndex & ":30"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":45 am"
                     
                     If intIndex = 0 Then
                     	strTime = "12:45 am"
					 End If	                 
                     
                     strTimeValue = intIndex & ":45"
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
                <!--
                	 <option selected value='12:00'>Noon</option>
                -->
                <%
				   Else
                %>
                <!--
                	 <option value='12:00'>Noon</option>
                -->
                <%
				   End If



                   For intIndex = 0 To 11  
                     strTime = intIndex & ":00 pm"

                     If intIndex = 0 Then
                     	strTime = "Noon"
					 End If	                 
                     
                     strTimeValue = intIndex + 12 & ":00"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":15 pm"

                     If intIndex = 0 Then
                     	strTime = "12:15 pm"
					 End If	                 
                     
                     strTimeValue = intIndex + 12 & ":15"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":30 pm"

                     If intIndex = 0 Then
                     	strTime = "12:30 pm"
					 End If	                 
                     
                     strTimeValue = intIndex + 12 & ":30"
                     If strTimeValue = strCheckTime Then
                     	strTimeSelected = "selected"
                     Else
                     	strTimeSelected = ""
                     End If
                %>
                	 <option <%=strTimeSelected %> value='<%=strTimeValue %>'><%=strTime %></option>
                <%
                     strTime = intIndex & ":45 pm"
                     
                     If intIndex = 0 Then
                     	strTime = "12:45 pm"
					 End If	                 
                     
                     strTimeValue = intIndex + 12 & ":45"
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
                <td width="106">&nbsp;</td>
                <td width="45">
                &nbsp;</td>
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
                Length of Rent:</td>
                <td width="92">
                
                <select name="lor" style="width:90; text-align:right" size="1">
                
                <% If adoRS4.Fields("lor").Value = 0 Then %>
                <option value='0' selected>0 days</option>
                <% Else %>
                <option value='0'>0 days</option>
                <% End If %> 
                <% If adoRS4.Fields("lor").Value = 1 Then %>
                <option value='1' selected>1 day</option>
                <% Else %>
                <option value='1'>1 day</option>
                <% End If %> 
                <% If adoRS4.Fields("lor").Value = 2 Then %>
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
                <option  value='11' selected>11 days</option>
                <% Else %>
                <option  value='11'>11 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 12 Then %>
                <option  value='12' selected>12 days</option>
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
                <% If adoRS4.Fields("lor").Value = 33 Then %>
                <option  value='33' selected>33 days</option>
                <% Else %>
                <option  value='33'>33 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 60 Then %>
                <option  value='60' selected>60 days</option>
                <% Else %>
                <option  value='60'>60 days</option>
                <% End If %>
                <% If adoRS4.Fields("lor").Value = 90 Then %>
                <option  value='90' selected>90 days</option>
                <% Else %>
                <option  value='90'>90 days</option>
                <% End If %>

<!--
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
-->
                </select>
                
                </td>
                <td width="99">
                <!-- 
              Mileage 
              Charges :
    			--></td>
                <td width="130">
                <!--              
              <select name="mileage_charges" style="width:90" size="1">
              <option selected value="0">Display</option>
              <option value="1">Disregard</option>
              </select>
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
                <td width="97">Data Source:
                </td>
                <td width="306" colspan="3">
                <% Dim strSelectDataSource
                
                   strSelectDataSource = adoRS4.Fields("data_sources").Value
                %>
                
                <select name="data_source" style="width:325; text-align:right" onchange="showCodeOption(this, 'XXX')">
                <% 'NOTE: IMPORTANT - if you add or modify these special cases
                    '     you must update "search_profile_insert_car.asp" AND
                    '                     "search_request_insert_car.asp"
                
                   If strSelectDataSource = "EXP,VAC" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M01"
                	   
              	   ElseIf strSelectDataSource = "EXP,VMW" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M02"
                	   
              	   ElseIf strSelectDataSource = "EXS,VAC" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M03"

              	   ElseIf strSelectDataSource = "EXD,VAC" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M04"

              	   ElseIf strSelectDataSource = "EXP,VAD" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M05"

              	   ElseIf strSelectDataSource = "WSR,VZE" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M06"

              	   ElseIf strSelectDataSource = "WSG,VZE" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M07"

              	   ElseIf strSelectDataSource = "WSA,VZE" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M08"

              	   ElseIf strSelectDataSource = "EXP,VZI" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M09"

              	   ElseIf strSelectDataSource = "WSR,VAD" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M10"

              	   ElseIf strSelectDataSource = "CRX,VET" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M11"

              	   ElseIf strSelectDataSource = "WSR,VZI,VZD,VET" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M12"

              	   ElseIf strSelectDataSource = "WSR,VEZ" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M13"

              	   ElseIf strSelectDataSource = "FFR,VET,VZE,VZT,VZR,VAL,VZL" Then
                	   Rem We have to translate the multi-site ones for selection on the site
                	   strSelectDataSource = "M14"

                   End If
                %>
                <% While adoRS1.EOF = False %>
                	<% If strSelectDataSource = adoRS1.Fields("data_source").Value Then %>
                		<option selected value="<%=adoRS1.Fields("data_source").Value %>">
                		<%=adoRS1.Fields("name").Value %></option>
                	<% ElseIf (strSelectDataSource = "") And (adoRS1.Fields("data_source").Value = "TRV") Then %>
                		<option selected value="<%=adoRS1.Fields("data_source").Value %>">
                		<%=adoRS1.Fields("name").Value %></option>
                	<% Else %>
                		<option value="<%=adoRS1.Fields("data_source").Value %>"><%=adoRS1.Fields("name").Value %>
                		</option>
                	<% End If %> 
                <% 
					 adoRS1.MoveNext
				   Wend
				   Set adoRS1 = Nothing
				%></select> </td>
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
                    <td width="473" colspan="4">
                    <div align="center">
                      (use buttons to move between selected/unselected)
                      <br>
                      Unselected types&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selected types&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                    </div>
                    </td>
                  </tr>
                  <tr>
                    <td width="211">
                      
                      <select name="unselected_car_types" size="5" style="width:200;"  multiple>
                      <% Dim CarTypes                                        %> 
                      <% Dim CarClasses                                      %> 
                      <% CarTypes = adoRS4.Fields("shop_car_type_cds").Value %>
                      <% While adoRS5.EOF = False                            %> 
                      <% If (InStr(CarTypes, adoRS5.Fields("car_type_cd").Value) = 0) Or (CarTypes = "") Then %>
                      <option value="<%=adoRS5.Fields("car_type_cd").Value %>"><%=adoRS5.Fields("class_desc").Value %>
                      </option>
                      <% End If %> 
                      <% adoRS5.MoveNext %> 
                      <% Wend %>
					  </select></td>
                    <td width="31">
                      
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_car_types, document.search_criteria.selected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_right.GIF" width="24" height="22" alt="Add the highlighted car types"    ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_car_types, document.search_criteria.selected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_right_all.GIF" width="24" height="22"  alt="All all the car types"  ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_car_types, document.search_criteria.unselected_car_types, false );update_car_type_list();return false;">
                      <img border="0" src="images/move_left.GIF" width="24" height="22"  alt="Remove the highlighted car types"  ></a><br>
                      
    
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_car_types, document.search_criteria.unselected_car_types, true );update_car_type_list();return false;">
                      <img border="0" src="images/move_left_all.GIF" width="24" height="22"  alt="Remove all the car types"  ></a></td>                    
                    <td width="116">
                    <font color="#080000">
                    
                    <select name="selected_car_types" size="5" style="width:200;" multiple>
                      <% Dim CarTypeArray                                    %> 
                      <% CarTypeArray = Split(CarTypes, ",")                 %>
                      <% For intIndex = 0 To UBound(CarTypeArray)            %>
                      <option value="<%=CarTypeArray(intIndex) %>"><%=CarTypeArray(intIndex) %>
                      </option>
                      <% Next %></select></font></td>
                    <td width="115">
                    <a href="javascript:void(0)" onclick="moveOptionUp(document.search_criteria.selected_car_types);update_car_type_list();return false;">
                      <img border="0" src="images/up_button.GIF" width="24" height="24"  alt="Click this up arrow to modify the display order of the car types" ></a><br>
                    <font color="#080000">
                    <a href="javascript:void(0)" onclick="moveOptionDown(document.search_criteria.selected_car_types);update_car_type_list();return false;">
                      <img border="0" src="images/down_button.GIF" width="24" height="24"  alt="Click this down arrow to modify the display order of the car types" ></a></font></td>
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
                    <td width="475" colspan="4">
                      <div align="center">
                        (use buttons to move between selected/unselected)
 
                        <br>
                        Unselected companies&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selected companies&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 
                      </div>
                 
                    </td>
                  </tr>
                  <tr>
                    <td width="209">
                    <select name="unselected_companies" size="5" style="width:200;" multiple>
                      <% Dim Companies                                       %> 
                      <% Dim SelectedCompanies							     %>
                      <% Companies = adoRS4.Fields("vend_cds").Value         %>
                      <% SelectedCompanies = "," & Companies & ","			 %> 
                      <% While adoRS2.EOF = False                            %>
                      <% If ((InStr("," & Companies & ",", "," & adoRS2.Fields("vendor_cd").Value & ",") = 0) Or (Companies = "")) And (adoRS2.Fields("vendor_cd").Value <> strSelfVendCd) Then %>
                      	<option value="<%=adoRS2.Fields("vendor_cd").Value %>"><%=adoRS2.Fields("vendor_name").Value %></option>
                      <% Else 		                    
		                    'SelectedCompanies = SelectedCompanies & TRIM(adoRS2.Fields("vendor_cd").Value) & "," & adoRS2.Fields("vendor_name").Value & "," 
		                    SelectedCompanies = Replace(SelectedCompanies, "," & adoRS2.Fields("vendor_cd").Value & ",", "," & TRIM(adoRS2.Fields("vendor_cd").Value) & "," & adoRS2.Fields("vendor_name").Value & ",")
						
						   End If 
						   adoRS2.MoveNext 
					   Wend 
					   Set adoRS2 = Nothing %>
					</select>
                      <% SelectedCompanies = Mid(SelectedCompanies,2,Len(SelectedCompanies) - 2)%>
					</td>
                    <td width="31">
                      
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_companies, document.search_criteria.selected_companies, false );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_right.GIF" width="24" height="22"  alt="Add the highlighted companies types"   ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.unselected_companies, document.search_criteria.selected_companies, true );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_right_all.GIF" width="24" height="22"  alt="Add all the companies"   ></a><br>
                      
    
                  <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_companies, document.search_criteria.unselected_companies, false );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_left.GIF" width="24" height="22"  alt="Remove the highlighted companies"   ></a><br>
                      
    
                    <a href="javascript:void(0)" onclick="moveDualList( document.search_criteria.selected_companies, document.search_criteria.unselected_companies, true );update_company_list(true, '');return false;">
                      <img border="0" src="images/move_left_all.GIF" width="24" height="22"  alt="Remove all the companies"   ></a></td>
                    <td width="114">
                      
                      <select name="selected_companies" size="5" style="width:200;" multiple>
                      <% Dim CompanyArray                                            %> 
                      <% Dim CompanyCodeList                                         %>
                      <% CompanyArray = Split(SelectedCompanies, ",")                %>
                      <% For intIndex = 0 To UBound(CompanyArray) - 1 Step 2         %>
                        <option value="<%=CompanyArray(intIndex) %>"><%=CompanyArray(intIndex + 1) %></option>
                        <% CompanyCodeList = CompanyCodeList & CompanyArray(intIndex) & ","  %>                                        
                      <% Next %>
                      </select></td>
                    <td width="114">
                      <a href="javascript:void(0)" onclick="moveOptionUp(document.search_criteria.selected_companies);update_company_list(true, '');return false;">
                      <img border="0" src="images/up_button.GIF" width="24" height="24" alt="Click this up arrow to modify the order" ></a><br>
                    <font color="#080000">
                    <a href="javascript:void(0)" onclick="moveOptionDown(document.search_criteria.selected_companies);update_company_list(true, '');return false;">
                      <img border="0" src="images/down_button.GIF" width="24" height="24"  alt="Click this down arrow to modify the order" ></a></font></td>
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
                
                <img src="images/ti_display_options.gif" width="162" height="25"></td>
                <td width="79">
                
                Comparison 
                Company : </td>
                <td width="209">
                
                <select name="highlighted_company" style="width:200" size="1">
                <% strHighlightVendor = adoRS4.Fields("highlight_vendor").Value %>
                <% If Len(SelectedCompanies) = 0 Then                           %>
	                <option selected>companies must be selected</option>
				<% Else														    %>
                <option >companies must be selected</option>
                <% For intIndex = 0 To UBound(CompanyArray) - 1 Step 2          %>
                	<% If (strHighlightVendor = CompanyArray(intIndex)) Then %>
		                <option selected value="<%=CompanyArray(intIndex) %>"><%=CompanyArray(intIndex + 1) %></option>
					<% Else %>
		                <option value="<%=CompanyArray(intIndex) %>"><%=CompanyArray(intIndex + 1) %></option>
					<% End If %>
                <% Next %>
                <% End If                                                      %>
                </select></td>
                <td width="28">
                <!--               
              <input type="checkbox" name="rate_drilldown" value="enable">
--></td>
                <td width="249">
                <!--              
              Enable 
              Rate Drill Down 
--></td>
              </tr>
            </table>
            <table width="745" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td background="images/alt_color.gif">
                <table width="745" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                  <tr valign="bottom">
                    <td width="161" >
                    <img src="images/separator.gif" width="30" height="15" alt=""></td>
                    <td >
                    <font size="2" color="#080000"><b>Note</b>: 
                    The comparison company will automatically be the first company 
                    displayed on reports</font></td>
                  </tr>
                </table>
                <table width="734" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                  <tr valign="bottom">
                    <td width="162" height="0">
                    
                    <img src="images/separator.gif" width="162" height="20"></td>
                    <td width="96" height="0">
                    
                    Rates Shown: </td>
                    <td width="247" height="0">
                    
                    <select name="display_rate_type" style="width:225" >
                    <% Select Case adoRS4.Fields("display_rate_type").Value %>
                    <%    Case 2   %>
                    <option value="1">Base rate amt (daily, wkly, etc)</option>
                    <option value="2" selected>Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <%    Case 3   %>
                    <option value="1">Base rate amt (daily, wkly, etc)</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3" selected >Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <%    Case 4   %>
                    <option value="1">Base rate amt (daily, wkly, etc)</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4" selected >Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <%    Case 5   %>
                    <option value="1">Base rate amt (daily, wkly, etc)</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5" selected >Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <%    Case 6   %>
                    <option value="1">Base rate amt (daily, wkly, etc)</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6" selected>Rate amt/Total price/Extra day</option>
                    <%    Case Else   %>
                    <option value="1" selected>Base rate amt (daily, wkly, etc)</option>
                    <option value="2">Total rate amount w/o fees</option>
                    <option value="3">Total price w/ fees</option>
                    <option value="4">Rate amt &amp; Total price</option>
                    <option value="5">Rate amt/Total price/Drop Chg</option>
                    <option value="6">Rate amt/Total price/Extra day</option>
                    <% End Select %>
                    </select> </td>
                    <td width="65" height="0">
                    
                    &nbsp;</td>
                    <td width="164" height="0">
                    
                    
                    &nbsp;</td>
                  </tr>
                  <tr valign="bottom">
                    <td width="162" height="0">
                    
                    &nbsp;</td>
                    <td width="96" height="0">
                    
                    <font color="#080000">
                    
                    Currency: </font>
					  </td>
                    <td width="247" height="0">
                    
                    <font color="#080000">
                    <select name="display_currency" style="width:239px" size="1">
                      <% Dim Currencies                                              %> 
                      <% Dim SelectedCurrency							             %>
                      <% SelectedCurrency = adoRS4.Fields("shop_currency_cd").Value  %>
                      <% While adoRSCurrency.EOF = False                             %>		
                      <%   If (adoRSCurrency.Fields("CurrencyCode").Value <> SelectedCurrency ) Then %>
						   <option value="<%=adoRSCurrency.Fields("CurrencyCode").Value %>"><%=adoRSCurrency.Fields("Country").Value %></option>
					  <%   Else %>	
                           <option selected="selected" value="<%=adoRSCurrency.Fields("CurrencyCode").Value %>"><%=adoRSCurrency.Fields("Country").Value %></option>
					  <%   End If
					  
	 					   adoRSCurrency.MoveNext 
					   	 Wend 
					     Set adoRSCurrency = Nothing 
					     
					  %>                      
					</select>
					</font>
					</td>
                    <td width="65" height="0">
                    
                    &nbsp;</td>
                    <td width="164" height="0">
                    
                    
                    &nbsp;</td>
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
                
                <img src="images/ti_actions.gif" width="162" height="25"></td>
                <td width="148">&nbsp;</td>
                <td width="423">
                &nbsp;</td>
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
                        <input type="radio" value="html" checked name="output">HTML 
                        Presentation</p>
                        </td>
                        <td width="188">
                        <select name="html_style" style="width:175" size="1" >
                        <% If adoRS4.Fields("output_style").Value = 2 Then %>
                        <option value="1">Grp by city and car class</option>
                        <option selected value="2">Grp by city and date</option>
                        <option value="3">Grp by custom car type groups</option>
                        <% ElseIf adoRS4.Fields("output_style").Value = 3 Then %>
                        <option value="1">Grp by city and car class</option>
                        <option value="2">Grp by city and date</option>
                        <option selected value="3">Grp by custom car type groups</option>
                        <% Else %>
                        <option selected value="1">Grp by city and car class</option>
                        <option value="2">Grp by city and date</option>
                        <option value="3">Grp by custom car type groups</option>
                        <% End If %>
                        
                        </select></td>
                      </tr>
                    </table>
                    <table width="546" border="0" cellpadding="2" cellspacing="0">
                      <tr valign="bottom">
                        <td style="width: 162px">
                        <img src="images/separator.gif" width="162" height="20"></td>
                        <td width="184">
                        <input type="radio" value="ftp" name="output" >Save 
                        to FTP server</td>
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
                        
                        <input type="radio" value="xls" name="output" disabled >Excel XLS 
                        Presentation</td>
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
                <table width="712" border="0" cellpadding="2" cellspacing="0">
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td>
                    
                    &nbsp;</td>
                    <td>
                    
                  
                                       &nbsp;</td>
                  </tr>
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td>
                    
                    Action : </td>
                    <td>
                    
                  
                                       <select name="search_action" style="width:225" size="1">
                    
                    <% If adoRS4.Fields("action_id").Value = 1 Then %>
                    	<option selected value="1">Search only (no emails)</option>
					<% Else %>
                    	<option value="1">Search only (no emails)</option>
                    <% End If %>
					
                    <% If adoRS4.Fields("action_id").Value = 2 Then %>
	                    <option selected value="2">Email rate &amp; suggestion rpt notices</option>
					<% Else %>
	                    <option value="2">Email rate &amp; suggestion rpt notices</option>
                    <% End If %>

                   <% If adoRS4.Fields("action_id").Value = 7 Then %>
	                    <option selected value="7">Email rate report notice only</option>
					<% Else %>
	                    <option value="7">Email rate report notice only</option>
                    <% End If %>

                    <% If adoRS4.Fields("action_id").Value = 3 Then %>
	                    <option selected value="3">Email suggestion rpt notice only</option>
					<% Else %>
	                    <option value="3">Email suggestion rpt notice only</option>
                    <% End If %>
<!--
                    <% If adoRS4.Fields("action_id").Value = 4 Then %>
	                    <option selected value="4">Email rate and alert reports</option>
					<% Else %>
	                    <option value="4">Email rate and alert reports</option>
                    <% End If %>
				
                    <% If adoRS4.Fields("action_id").Value = 5 Then %>
	                    <option selected value="5">Email alert report only</option>
					<% Else %>
	                    <option value="5">Email alert report only</option>
                    <% End If %>
-->	
                    <% If adoRS4.Fields("action_id").Value = 6 Then %>
	                    <option selected value="6">Email change reciept only</option>
					<% Else %>
	                    <option value="6">Email change receipt only</option>
                    <% End If %>

					
                    </select>

                    (all options search)</td>
                  </tr>
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td>
                    &nbsp;&nbsp;&nbsp; 
                    Report E-mail Address:</td>
                    <td>
                    
                    <input type="text" name="recipient_address" size="26" style="width:175" value="<%=adoRS4.Fields("email_address").Value %>"> 
                    (separate with commas)</td>
                  </tr>
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td valign="middle">
                    
                    &nbsp;&nbsp;&nbsp; Alert E-mail Address:</td>
                    <td>
                    
                    <input type="text" name="alert_address" size="26" style="width:175" value="<%=adoRS4.Fields("alert_address").Value %>"> 
                    (separate with commas)</td>
                  </tr>
                  <tr valign="bottom">
                    <td>&nbsp;</td>
                    <td valign="top">
                    &nbsp;</td>
                    <td>
                    
                    (leave blank if same address as report e-mail address)</td>
                  </tr>
                  <tr valign="bottom">
                    <td>
                    <img src="images/separator.gif" width="162" height="20" alt=""></td>
                    <%
						strShopPOS = adoRS4.Fields("shop_pos_cd").Value
					%>
                    <td>Shop POS:</td>
                    <td><select name="pos" style="width:294px">
                    
                    <% Dim Countries                                               %> 
                    <% Dim SelectedCountry                                                                                                                     %>
                    <% SelectedCountry = adoRS4.Fields("shop_pos_cd").Value        %>
                    <% While adoRSCountry.EOF = False                               %>                             
                    <%   If (adoRSCountry.Fields("country_cd").Value <> SelectedCountry ) Then %>
                            <option value="<%=adoRSCountry.Fields("country_cd").Value %>"><%=adoRSCountry.Fields("country_name").Value %></option>
                    <%   Else %>     
                            <option selected="selected" value="<%=adoRSCountry.Fields("country_cd").Value %>"><%=adoRSCountry.Fields("country_name").Value %></option>
                    <%   End If
                                                                               
                         adoRSCountry.MoveNext 
                       Wend 
                       Set adoRSCountry = Nothing 
                                                       
                    %>
                    
                                       
<!--                    <option value="UK">United Kingdom</option>
                    <option selected value="US">United States</option>
                    <option value="FR">France</option>
                    <option value="CA">Canada</option>
                    <option value="AU">Australia</option>
                    <option value="NZ">New Zealand</option> 
-->                    
                    
                    </select>
                    </td>
                  </tr>
                  </table>

                  <div id="discount_area" name="discount_area" style="display:block">
                  <table width="712" border="0" cellpadding="2" cellspacing="0" id="discount_codes" name="discount_codes">
                  <tr valign="bottom">
                    <td style="width: 162px">
                    <img src="images/separator.gif" width="162" height="20"></td>
                    <td width="162">
                    Discount Code(s):</td>
                    <td width="114">
                    
					<i>Discount Code</i></td>
                    <td width="247">
                    
					<i>Company</i></td>
                  </tr>
                </table>

				</div>
                <table width="713" border="0" cellpadding="2" cellspacing="0">
                  <tr valign="bottom">
                    <td style="width: 159px">
                    <img src="images/separator.gif" width="162" height="20"></td>
                    <td width="162">Scheduled Search Time: </td>
                    <td width="365">
                    
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
                	 <option selected value='00:00'>Midnight (12:00 am)</option>
                
                <% Else %>
                	 <option value='00:00'>Midnight (12:00 am)</option>
                
                <% End If %>

                <% If InStr(1, strCheckTime, "00:15") > 0 Then %>
                	 <option selected value='00:15'>12:15 am</option>
                
                <% Else %>
                	 <option value='00:15'>12:15 am</option>
                
                <% End If %>

                <% If InStr(1, strCheckTime, "00:30") > 0 Then %>
                	 <option selected value='00:30'>12:30 am</option>
                
                <% Else %>
                	 <option value='00:30'>12:30 am</option>
                
                <% End If %>

                <% If InStr(1, strCheckTime, "00:45") > 0 Then %>
                	 <option selected value='00:45'>12:45 am</option>
                
                <% Else %>
                	 <option value='00:45'>12:45 am</option>
                
                <% End If %>

                
                <%
                

                   For intIndex = 1 To 11  
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 am"
                     	Case 1
	                       strTime = intIndex & ":15 am"
                     	Case 2
	                       strTime = intIndex & ":30 am"
                     	Case 3
	                       strTime = intIndex & ":45 am"
					 End Select	                       
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
                   Next   


				   If strCheckTime = "12:00" Then	
                %>
                	 <option selected value='12:00'>Noon (12:00 pm)</option>
                <%
				   Else
                %>
                	 <option value='12:00'>Noon (12:00 pm)</option>
                <%
				   End If


				   If strCheckTime = "12:15" Then	
                %>
                	 <option selected value='12:15'>12:15 pm</option>
                <%
				   Else
                %>
                	 <option value='12:15'>12:15 pm</option>
                <%
				   End If

				   If strCheckTime = "12:30" Then	
                %>
                	 <option selected value='12:30'>12:30 pm</option>
                <%
				   Else
                %>
                	 <option value='12:30'>12:30 pm</option>
                <%
				   End If

				   If strCheckTime = "12:45" Then	
                %>
                	 <option selected value='12:45'>12:45 pm</option>
                <%
				   Else
                %>
                	 <option value='12:45'>12:45 pm</option>
                <%
				   End If





                   For intIndex = 1 To 11 
                     For intMinuteIndex = 0 To 3 
                     Select Case intMinuteIndex
                     	Case 0
	                       strTime = intIndex & ":00 pm"
                     	Case 1
	                       strTime = intIndex & ":15 pm"
                     	Case 2
	                       strTime = intIndex & ":30 pm"
                     	Case 3
	                       strTime = intIndex & ":45 pm"
					 End Select	                       
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
                   Next   
                %>
                    </select> (local time).</td>
                  </tr>
                </table>
                <table width="743" border="0" cellpadding="2" cellspacing="0">
                  <tr valign="bottom">
                    <td style="width: 157px">
                    <img src="images/separator.gif" width="162" height="20"></td>
                    <td width="162">
                    &nbsp;&nbsp;&nbsp; 
                    Days of Week:</td>
                    <td width="395">
                    <table width="381" border="0" cellpadding="2" cellspacing="0" background="images/alt_color.gif">
                      <tr valign="bottom">
                        <td width="20">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "1") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="1" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="1"> 
		                <% End If %>
                        </td>
                        <td width="29">
                        Sun
                        </td>
                        <td width="20">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "2") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="2" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="2"> 
		                <% End If %>
                        </td>
                        <td width="25">
                        Mon
                        </td>
                        <td width="20">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "3") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="3" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="3"> 
		                <% End If %>
                        </td>
                        <td width="31">
                        Tue </td>
                        <td width="20">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "4") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="4" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="4"> 
		                <% End If %>
                        </td>
                        <td width="25">
                        Wed
                        </td>
                        <td width="20">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "5") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="5" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="5"> 
		                <% End If %>
                        </td>
                        <td width="19">
                        Thu
                        </td>
                        <td width="20">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "6") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="6" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="6"> 
		                <% End If %>
                       </td>
                        <td width="20">
                        Fri</td>
                        <td width="20">
                        
		                <% If Instr(adoRS4.Fields("schedule_dow_list").Value, "7") Then %>
		                	<input type="checkbox" name="schedule_dow_list" value="7" checked> 
		                <% Else %>
		                	<input type="checkbox" name="schedule_dow_list" value="7"> 
		                <% End If %>
                       </td>
                        <td width="20">
                        <font size="2" color="#080000">Sat</font> </td>
                      </tr>
                    </table>
                    </td>
                  </tr>
                  <tr valign="bottom">
                    <td style="width: 157px">
                    &nbsp;</td>
                    <td width="162">
                    &nbsp;</td>
                    <td width="395">
                    &nbsp;</td>
                  </tr>
                  <tr valign="bottom">
                    <td style="width: 157px">&nbsp;</td>
                    <td width="162">Advanced Scheduler</td>
                    <td width="395"><span class="style2">(beta)</span></td>
                  </tr>
                  <tr valign="bottom">
                    <td style="width: 157px">&nbsp;</td>
                    <td width="162">&nbsp;&nbsp; Selected Schedule:&nbsp;</td>
                    <td width="395">
                    <!-- <input name="search_schedules" type="button" id="search_schedules" value="Search Schedules" class="button" onclick="window.open('transfer.asp?goto=ProfileSearchScheduleA.aspx','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=yes,width=820,height=600'); return false;"> -->
                    <select name="search_schedules" id="search_schedules" style="width: 298px" > 
                    <% If ProfileID = 0 Then %>
					<option value="0" selected="selected"  >(none selected)</option>
					<% While adoRS6.EOF = False %>
                      	<option value="<%=adoRS6.Fields("schedule_grp_id").Value %>"><%=adoRS6.Fields("schedule_grp_desc").Value %></option>
					<%  adoRS6.MoveNext 
					   Wend %>
				

					<% Else %>
					<option value="0" >(none selected)</option>
					<% While adoRS6.EOF = False %>
					  <% SchdProfileID = adoRS6.Fields("profile_id").Value %>
                      <% If (adoRS6.Fields("profile_id").Value = ProfileID) Then %>
                      	<option selected="selected" value="<%=adoRS6.Fields("schedule_grp_id").Value %>"><%=adoRS6.Fields("schedule_grp_desc").Value %></option>
                      <% Else %>
                      	<option value="<%=adoRS6.Fields("schedule_grp_id").Value %>"><%=adoRS6.Fields("schedule_grp_desc").Value %></option>
						
					  <% End If 
						   adoRS6.MoveNext 
					   Wend 
					
					End If   
					   
					Set adoRS6 = Nothing %>
					</select>
					</td>
                  </tr>
                  <tr valign="bottom">
                    <td style="width: 157px">&nbsp;</td>
                    <td width="162">&nbsp;</td>
                    <td width="395">&nbsp;<a href="javascript:void()" onclick="window.open('transfer.asp?goto=ProfileSearchScheduleAll.aspx','manage_schedule','toolbar=no,status=no,scrollbars=yes,resizable=yes,width=820,height=600'); return false;" >Advanced Scheduler</a>
					
            		</td>
                  </tr>
                  <tr valign="bottom">
                    <td style="width: 157px">&nbsp;</td>
                    <td width="561" colspan="2">
                    &nbsp;</td>
                  </tr>
                  <tr valign="bottom">
                    <td style="width: 157px">&nbsp;</td>
                    <td width="561" colspan="2">
                    <input name="submit" type="submit" id="submit_request" value="Submit Request" class="rh_button" onclick="return confirmSubmit('searchnow')"></td>
                  </tr>
                  <tr valign="bottom">
                    <td style="width: 157px"></td>
                    <td width="561" colspan="2">&nbsp;</td>
                  </tr>
                </table>
                </td>
              </tr>
            </table></font>
            </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
  
  <%   If strUserId = 1 Then	%>
   
              <p align="center">Debug section:<br>
              <input type="text" name="cities_list"  value="<%=adoRS4.Fields("city_cd").Value%>">
              <input type="text" name="car_type_list" value="<%=adoRS4.Fields("shop_car_type_cds").Value %>">
              <input type="text" name="company_list" value="<%=CompanyCodeList %>" > 
              
              <br>
              <input type="checkbox" name="debug" value="true" checked>Debug 
              (don't alter the database, just display the results)<br>
              strCheckTime=<%=strCheckTime  %><br></p>
  <%   Else	%>
              <input type="hidden" name="company_list" value="<%=CompanyCodeList %>"><br>
              <input type="hidden" name="car_type_list" size="20" value="<%=adoRS4.Fields("shop_car_type_cds").Value %>"><br>
  			  <input type="hidden" name="debug" value="false" > 
			  <input type="hidden" name="cities_list" value="<%=adoRS4.Fields("city_cd").Value%>">
  <%   End If	%>
        
</form>
<script type="text/javascript"  language="JavaScript">
    update_company_list(false, '<%=adoRS4.Fields("extra_criteria").value %>');
    // Now show the discount codes if the user loaded a profile with them
    showCodeOption('not an obj', '<%=strSelectDataSource %>');
</script>
<div id="calbox" class="calboxoff"></div>	
<%
	If err.number = 66660 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>An error cccured while loading the profile information</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<%	
	Rem Clean up
	Set adoCmd  = Nothing 
	Set adoCmd1 = Nothing 
	Set adoCmd2 = Nothing 
	Set adoCmd3 = Nothing 
	Set adoCmd4 = Nothing
	Set adoCmd5 = Nothing 	


	Set adoRS  = Nothing 
	Set adoRS1 = Nothing 
	Set adoRS2 = Nothing 
	Set adoRS3 = Nothing 
	Set adoRS4 = Nothing
	Set adoRS5 = Nothing 
	Set adoRS6 = Nothing

	
%>