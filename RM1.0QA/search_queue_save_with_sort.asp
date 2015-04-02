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
	Dim adoPrices
	Dim intIndex
	Dim blnExit
	Dim intReportRequestID
	Dim intErrorCount
	Dim intTimeoutCount
	Dim strCarType 
	Dim intResults
	Dim intPrice
	Rem we have no clue how many, so cross your fingers
	Dim varCarTypes()
	Dim varDataSources()
	Dim varDates()

	strClientUserid = Request.Form("userid")
	strCity = Request.Form("city")
	strCarType = Request.Form("car_type")
	strCompany = Request.Form("company")
	
	strSearched = False
	
	If (strClientUserid = "") And (strCity = "") And (strCarType  = "") And (strCompany = "") Then
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select1"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Refresh 

		adoCmd.Parameters("@user_id").Value = Session("user_id")
		adoCmd.Parameters("@client_userid").Value = Null
		adoCmd.Parameters("@city_cd").Value = Null
		adoCmd.Parameters("@shop_car_type_cds").Value = Null
		adoCmd.Parameters("@vendor_cd").Value = Null 
		
		Set adoRS = adoCmd.Execute
	
		strSearched = True
 
	Else
	
		strConn = Session("pro_con")
	
  		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_shop_request_select1"
		adoCmd.CommandType = 4

		adoCmd.Parameters.Refresh 

		If Trim(strClientUserid) <> "" Then
			adoCmd.Parameters("@client_userid").Value = strClientUserid 
		Else
			adoCmd.Parameters("@client_userid").Value = Null
		End If


		If Trim(strCity) <> "" Then
			adoCmd.Parameters("@city_cd").Value = strCity 
		Else
			adoCmd.Parameters("@city_cd").Value = Null
		End If


		If Trim(strCarType) <> "" Then
			adoCmd.Parameters("@shop_car_type_cds").Value = strCarType 
		Else
			adoCmd.Parameters("@shop_car_type_cds").Value = Null
		End If
	

		If Trim(strCompany) <> "" Then
			adoCmd.Parameters("@vendor_cd").Value = strCompany 
		Else
			adoCmd.Parameters("@vendor_cd").Value = Null 
		End If

		Set adoRS = adoCmd.Execute
		'Set adoRS1 = adoCmd.Execute
	
		strSearched = True
		

	
	
	End If




	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Queue</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script src="inc/sitewide.js" language="javascript"></script>
<SCRIPT LANGUAGE="JavaScript">
function myObject() {
    for (var i = 0; i<myObject.arguments.length; i++)
        this['n' + i] = myObject.arguments[i];
}

var objectArrayIndex = 0;
var myObjectArray = new Array();


function showObjectArray(object,length) {
    var output = '<CENTER><TABLE BORDER=1>';
    output += '<TR>';
    for (var j=0; j<width; j++)
         output += '<TH><A HREF="search_queue.asp?n' + j + '">' + eval('object[0].n'+j) + '<\/A><\/TH>';
    output += '<\/TR>';
    for (var i=1; i<length; i++) {
        output += '<TR>';
        for (var j=0; j<width; j++)
            output += '<TD>' + eval('object[i].n'+j) + '<\/TD>';
        output += '<\/TR>';
    }
    output += '<\/TABLE><\/CENTER>';
    document.write(output);
}

function myObjectBubbleSort(arrayName,length,property) {
    for (var i=1; i<(length-1); i++)
        for (var j=i+1; j<length; j++)
            if (eval('arrayName[j].' + property + '<arrayName[i].' + property)) {
                var dummy = arrayName[i];
                arrayName[i] = arrayName[j];
                arrayName[j] = dummy;
            }
}
</script>
<%
response.write "<script language='JavaScript' type='text/JavaScript'> " & vbcrlf
'response.write " <!-- " & vbcrlf
'response.write " var objectArrayIndex = 0; " & vbcrlf
'response.write " var myObjectArray = new Array();" & vbcrlf
response.write " var width = 12; " & vbcrlf
response.write " myObjectArray[objectArrayIndex++] = new myObject('Search ID','Search Status','User','Profile', 'Action', 'Search Units', 'Rate Units Expected', 'Rate Units Complete', 'Pickup City', 'First Rental Date', 'Last Rental Date', 'Car Types', 'Companies' );" & vbcrlf
While adoRS.EOF = False
response.write " myObjectArray[objectArrayIndex++] = new myObject('" & adoRS.Fields("shop_request_id").Value & "'"
response.write ", '" & adoRS.Fields("request_status").Value & "'"
response.write ", '" & adoRS.Fields("client_userid").Value & "'"
response.write ", '[none]'"
response.write ", 'Display'"
response.write ", '" & adoRS.Fields("work_units").Value & "'"
response.write ", '" & adoRS.Fields("work_units").Value & "'"
response.write ", '" & adoRS.Fields("work_units_complete").Value & "'"
response.write ", '" & adoRS.Fields("city_cd").Value & "'"
response.write ", '" & adoRS.Fields("begin_arv_dt").Value & "'"
response.write ", '" & adoRS.Fields("end_arv_dt").Value & "'"
response.write ", '" & adoRS.Fields("shop_car_type_cds").Value & "'"
response.write ", '" & adoRS.Fields("vend_cd").Value & "'"
response.write "); " &  vbcrlf
adoRS.MoveNext
Wend
'response.write " --> " & vbcrlf
response.write " </script> " & vbcrlf
%>

<script language="JavaScript" type="text/JavaScript">

var sortProperty = window.location.search.substring(1);

if (sortProperty.length != 0) 
    myObjectBubbleSort(myObjectArray,objectArrayIndex,sortProperty);

showObjectArray(myObjectArray,objectArrayIndex);
//-->
</script>
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
//-->
</script>
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
    <table width="400" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/b_left.jpg" width="62" height="32"></td>
        <td>
        <a href="search_profiles_car.asp" onmouseover="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td>
        <td>
        <a href="search_queue_save_with_sort.asp" onmouseover="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td>
        <td>
        <a href="search_criteria_car.asp" onmouseover="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('al','','images/b_alert_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('us','','images/b_user_on.gif',1)" onmouseout="MM_swapImgRestore()">
        <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td>
        <td>
        <a href="javascript:not_enabled()" onmouseover="MM_swapImage('sy','','images/b_system_on.gif',1)" onmouseout="MM_swapImgRestore()">
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
                <td>
                <div align="right">
                  <font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
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
        <td><img src="images/h_search_que.gif" width="368" height="31"></td>
        <td><img src="images/h_right.gif" width="402" height="31"></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
&nbsp;
<form method="POST" action="search_queue_save_with_sort.asp" name="search" class="search">
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
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
  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="1110" id="AutoNumber1" background="images/alt_color.gif">
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="278" height="18" colspan="2">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">To search enter 
      the last name, or a portion of. You may also optionally enter city, car type 
      and/or the car company.</font></td>
      <td width="699" height="18">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="26">&nbsp;</td>
      <td width="182"><img border="0" src="images/search.GIF"></td>
      <td width="87" height="26">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">User Name:
      </font></td>
      <td width="191" height="26">
      <input type="text" name="userid" size="20" value="<%=strClientUserid %>" onfocus="this.className='focus';cl(this,'<%=strClientUserid %>');" onblur="this.className='';fl(this,'<%=strClientUserid %>');">
      </td>
      <td width="699" height="26">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <input type="submit" value="  Display  " name="submit" class="rh_button"></font></td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">City:</font></td>
      <td width="191" height="22">
      <input type="text" name="city" size="20" value="<%=strCity %>" onfocus="this.className='focus';cl(this,'<%=strCity %>');" onblur="this.className='';fl(this,'<%=strstrCity %>');"></td>
      <td width="699" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Car Type:</font></td>
      <td width="191" height="22">
      <input type="text" name="car_type" size="20" value="<%=strCarType %>" onfocus="this.className='focus';cl(this,'<%=strCarType %>');" onblur="this.className='';fl(this,'<%=strCarType %>');"></td>
      <td width="699" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="22">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="22">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Company:</font></td>
      <td width="191" height="22">
      <input type="text" name="company" size="20" value="<%=strCompany %>" onfocus="this.className='focus';cl(this,'<%=strCompany %>');" onblur="this.className='';fl(this,'<%=strCompany %>');"></td>
      <td width="699" height="22">&nbsp;</td>
    </tr>
    <tr>
      <td width="19" height="18">&nbsp;</td>
      <td width="182">&nbsp;</td>
      <td width="87" height="18">&nbsp;</td>
      <td width="191" height="18">&nbsp;</td>
      <td width="699" height="18">&nbsp;</td>
    </tr>
  </table>
  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
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
</form>
<table width="1110" border="0" cellpadding="2" background="images/alt_color.gif" style="border-collapse: collapse" bordercolor="#111111">
  <tr valign="bottom">
    <td width="169">&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="search_profiles_car.asp">|&lt;</a>
    <a href="search_profiles_car.asp">&lt;</a> Page 1 of 1
    <a href="search_profiles_car.asp">&gt;</a> <a href="search_profiles_car.asp">&gt;|</a></font></td>
  </tr>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
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
<thead >
  <tr>
    <td align="left" valign="bottom" bgcolor="#879AA2" height="45" width="26">&nbsp;</td>
    <td class="profile_header" width="63" style="background-color: #E07D1A" height="45">
    Selected</td>
    <td class="profile_header" width="57" height="45">Search ID</td>
    <td class="profile_header" width="100" height="45">Search Status</td>
    <td class="profile_header" width="46" height="45">User</td>
    <td class="profile_header" width="58" height="45">Profile or [none]</td>
    <td class="profile_header" width="58" height="45">Action</td>
    <td class="profile_header" width="76" height="45">Search Units</td>
    <td class="profile_header" width="73" height="45">Rate Units Expected</td>
    <td class="profile_header" width="79" height="45">Rate Units Complete</td>
    <td class="profile_header" width="70" height="45">Pickup City</td>
    <td class="profile_header" width="72" height="45">First Rental Date</td>
    <td class="profile_header" width="82" height="45">Last Rental Date</td>
    <td class="profile_header" width="82" height="45">Car Types</td>
    <td class="profile_header" width="82" height="45">Companies</td>
  </tr>
  </thead> 
  <%
        
        Dim strClass
        Dim strOrange
        Dim intCount

		If strSearched = True Then

		While adoRS.EOF = False
		
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
    <td width="26" class="<%=strClass %>" height="20"><%=intCount  %></td>
    <td width="63" <%=strOrange%> align="center" height="20">
    <input type="radio" value="V1" name="selected"></td>
    <td width="57" class="<%=strClass %>" height="20">
    <a href="view_report.asp?ReportRequestID=<%=adoRS.Fields("shop_request_id").Value %>" target="_blank">
    <%=adoRS.Fields("shop_request_id").Value %></a></td>
    <td width="100" class="<%=strClass %>" height="20"><%=adoRS.Fields("request_status").Value %></td>
    <td width="46" class="<%=strClass %>" height="20"><%=adoRS.Fields("client_userid").Value %></td>
    <td width="58" class="<%=strClass %>" height="20">[none]</td>
    <td width="58" class="<%=strClass %>" height="20">Display</td>
    <td width="76" class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units").Value %></td>
    <td width="73" class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units").Value %></td>
    <td width="79" class="<%=strClass %>" height="20"><%=adoRS.Fields("work_units_complete").Value %></td>
    <td width="70" class="<%=strClass %>" height="20"><%=adoRS.Fields("city_cd").Value %></td>
    <td width="72" class="<%=strClass %>" height="20"><%=adoRS.Fields("begin_arv_dt").Value %></td>
    <td width="82" class="<%=strClass %>" height="20"><%=adoRS.Fields("end_arv_dt").Value %></td>
    <td width="82" class="<%=strClass %>" height="20"><%=adoRS.Fields("shop_car_type_cds").Value %></td>
    <% If adoRS.Fields("vend_cd").Value = "" Then %>
    <td width="82" class="<%=strClass %>" height="20">All</td>
    <% Else %>
    <td width="82" class="<%=strClass %>" height="20"><%=adoRS.Fields("vend_cd").Value %></td>
    <% End If %>
  </tr>
  <%
        
        	adoRS.MoveNext
        	
        Wend
        
   		adoRS.Close
		Set adoRS1 = Nothing
		Set adoCmd = Nothing

		Else

		%>
		
		
  <tr>
    <td width="26" class="profile_light" height="20"></td>
    <td width="63" bgcolor="#FDC677" align="center" height="20">
    <input type="radio" value="V1" name="selected"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="26" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="13" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
    <td width="1" class="profile_light" height="20"></td>
  </tr>
  <%

		End If
		        
        %>
</table>
<table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1110" height="4">
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
<p>&nbsp;| <a href="http://orion.mysymmetry.net/CARS/delete_alert.asp">Cancel</a> 
| <a href="http://orion.mysymmetry.net/CARS/copy_alert.asp">Delete</a> | </p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
