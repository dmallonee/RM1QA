<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% Response.Expires = -1  
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache" 
   
   	on error resume next

   	Server.ScriptTimeout = 30

	strUserId = Request.Cookies("rate-monitor.com")("user_id")
	
	strConn = Session("pro_con")
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)

	' City codes		
	Set adoRS1 = adoCmd.Execute
	' Return city codes
	Set adoRS2 = adoCmd.Execute
	
	adoCmd.CommandText = "car_type_select"

	' Car types		
	Set adoRS3 = adoCmd.Execute
	
	
	If Request("delete") = "true" Then
		Set adoCmd = CreateObject("ADODB.Command")
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_fleet_adjustment_delete"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_id",      3, 1, 0, Request("car_id"))
			
		adoCmd.Execute
		
	End If
	
	If Request("loc_cd") <> "" Then
	
		Set adoCmd = CreateObject("ADODB.Command")
	
		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_fleet_adjustment_insert"
		adoCmd.CommandType = adCmdStoredProc
	
		If IsDate(Request("return_dttm")) Then
			ReturnDate = Request("return_dttm")
		Else
			ReturnDate = Null
		End If

		If IsDate(Request("begin_dt")) Then
			BeginDate = Request("begin_dt")
		Else
			BeginDate = Null
		End If

		If IsDate(Request("end_dt")) Then
			EndDate = Request("end_dt")
		Else
			EndDate = Null
		End If


	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",           3, 1,  0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_desc",	    200, 1, 50, Request("car_desc"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd",     200, 1,  4, UCASE(Request("car_type_cd")))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@loc_cd", 		    200, 1,  6, UCASE(Request("loc_cd")))		
		adoCmd.Parameters.Append adoCmd.CreateParameter("@rtrn_loc_cd",     200, 1,  6, UCASE(Request("rtrn_loc_cd")))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@status_cd",      	  2, 1,  0, Request("status_cd"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@return_dttm",     135, 1,  0, ReturnDate)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_count",         3, 1,  0, Request("car_count"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@begin_dt",        135, 1,  0, BeginDate)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@end_dt",          135, 1,  0, EndDate)
					
		Set adoRS = adoCmd.Execute

	
	End If
		
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_fleet_adjustment_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",      3, 1, 0, strUserId)
				
	Set adoRS = adoCmd.Execute

	
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br/>"
	   response.write pad & "</b><br/>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br/>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br/>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br/>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br/>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br/><hr>"

	End If
  
  
%>    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Language" content="en-us" />
<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; RezCentral Fleet Adjustment</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
<link rel="stylesheet" type="text/css" href="inc/sitewide.css" />
<link rel="stylesheet" type="text/css" href="inc/rh_report.css" />
<script language="JavaScript" type="text/JavaScript" src="inc/sitewide.js" ></script>
<script language="javascript" type="text/javascript" src="inc/pupdate.js"></script>
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

<link rel="stylesheet" type="text/css" href="inc/rh_standard.css" />
<style type="text/css">
<!--
.profile_header_center { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style5 {
	text-align: center;
	font-size: medium;
}
.style7 {
	border-collapse: collapse;
	width: 750px;
	border-color:#111111;
	
	
}
.style10 {
	text-align: right;
	border-style: solid;
	border-width: 0;
}
.style11 {
	text-align: center;
}
.style12 {
	font-size: xx-small;
}
.style13 {
	text-align: right;
}
.style14 {
	font-size: xx-small;
	text-align: center;
}
-->
</style>
<base target="_self" />
  <style type="text/css"> 
      .LockOff { 
         display: none; 
         visibility: hidden; 
      } 

      .LockOn { 
         display: block; 
         visibility: visible; 
         position: absolute; 
         z-index: 999; 
         top: 0px; 
         left: 0px; 
         width: 105%; 
         height: 105%; 
         background-color: #ccc; 
         text-align: center; 
         padding-top: 20%; 
         filter: alpha(opacity=75); 
         opacity: 0.75; 
      } 
   .auto-style1 {
	  border-style: solid;
	  border-width: 0;
	  text-align:left;
	  
  }
   </style> 

   <script type="text/javascript"> 
      function skm_LockScreen(str) 
      { 
         var lock = document.getElementById('skm_LockPane'); 
         if (lock) 
            lock.className = 'LockOn'; 

         lock.innerHTML = str; 
      } 
   </script> 

</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif"><img src="images/top.jpg" width="770" height="91" /></td>
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
    <td background="images/med_bar_tile.gif"><img src="images/med_bar.gif" width="12" height="8" /></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td background="images/user_tile.gif">
<table width="100" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><img src="images/user_left.gif" width="580" height="31" /></td>
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
                      <td><img src="images/separator.gif" width="183" height="6" /></td>
                    </tr>
                  </table>
                </td>
                <td><img src="images/user_tile.gif" width="7" height="31" /></td>
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
        <td><img src="images/h_system.gif" width="368" height="31" /></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25" />
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0" /></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<p class="style11">&nbsp;&nbsp;&nbsp; <br/>
&nbsp;<img alt="RezCentral" src="images/rezcentral.jpg"  /><strong> 
</strong></p>
  <table border="0" style="width: 800px;" bordercolor="#FFFFFF"  align="center" class="style7">
    <tr>
      <td ><font size="2" face="Vendana, Arial, Helvetica, sans-serif">[<a href="rezcentral_tethering_20130715.asp">tether settings</a>]
[<a href="rezcentral_tethering_ow_20130715.asp">tethering one-way settings</a>] [<a href="system_queue_rezcentral.asp">queue 
status</a>] [<a href="rezcentral_update_status.asp">report status</a>] [<a href="RezCentralHeader.aspx?uid=<%=strUserId %>">Rate 
		Code Detail</a>] [<a href="https://rezcentral.tsdasp.net/WebRezClient/" target="_blank">login 
to RezCentral</a>] [<a href="rezcentral_blocks.asp">block settings</a>]&nbsp; <b>[fleet adjustment] </b>
</font></td>
     
    </tr>
    <tr>
      <td >&nbsp;</td>
     
    </tr>
    <tr>
      <td background="images/ruler.gif" height="4"></td>
     
    </tr>
  </table>
<p>&nbsp;</p><!-- 
	<p align="center">
		Peter - use this report for right now please =&gt;
	<a href="system_utilization_report.asp">utilization report</a>
	</p>
	--><form method="get" name="add_fleet_adjustment" >
<div class="style5">&nbsp;Current Fleet Adjustment Settings&nbsp;<br/><br/>
        <table border="0" cellpadding="0"  class="style7" align="center">
          <tr>
           <td class="auto-style1" colspan="11"><font size="2"><b>
           Directions:</b> To create a new fleet adjustment for RezCentral/TSD 
		   utilization, enter the 
			values in the fields at the bottom of the list. To delete an 
		   adjustment, 
			click the delete link to the right of the adjustment to delete. To 
			update, simply delete then recreate. Return branch should be the 
		   same as branch for local rentals. Required fields have an asterisk.<br/>&nbsp;</font><br/>
		   <font size="2"><strong>Note</strong>: Dates must be entered in 
		   MM/DD/YYYY format. A date is only required for a one-way fleet 
		   adjustment. Please use a status of zero (0) unless instructed to do 
		   otherwise. The begin and end dates are only required if the 
		   adjustement is desired for a limited time-frame, if so only enter the 
		   dates that you want that car type adjusted for.</font><br/>
           
           </td>
           
          </tr>
          </table>
          <table border="0" cellpadding="0" style="width: 850;" bordercolor="#111111" class="style7" align="center">
          <tr>
           	<td class="boxtitle" >&nbsp;</td>
           	<td class="style10" >&nbsp;</td>
			<td class="boxtitle" colspan="2" >
			&nbsp;</td>
            <td class="boxtitle" colspan="2">
			&nbsp;</td>
            <td class="boxtitle">&nbsp;</td>

            <td class="boxtitle">&nbsp;</td>

            <td class="boxtitle">&nbsp;</td>

            <td class="boxtitle">&nbsp;</td>

            <td class="boxtitle">&nbsp;</td>

		  </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"> 
          	</td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           <td  class="boxtitle"  style="height: 15px"></td>
           
           <td  class="boxtitle"  style="height: 15px"></td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px"></td>
           
          </tr>
          <tr>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">
			&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
           <td  class="boxtitle"  style="height: 15px">&nbsp;</td>
           
          </tr>
          <tr>
            <td  class="profile_header_center" >Branch*</td>
            <td  class="profile_header_center" >Rtrn Branch*</td>
            <td  class="profile_header_center" >Class Code*</td>
            <td  class="profile_header" >Note</td>
            <td  class="profile_header_center" >Status Code*</td>
            <td  class="profile_header_center" >Return Date</td>
            <td  class="profile_header_center" >Count*</td>
            <td  class="profile_header_center" >Begin</td>
            <td  class="profile_header_center" >End</td>
            <td  class="boxtitle" >&nbsp;</td>
            <td  class="boxtitle" >&nbsp;</td>
            <td  class="boxtitle" >&nbsp;</td>
          </tr>
          <% 	Dim intCount 	 %>
          <%    Dim bolLight     %>
          <%    Dim strClass     %>
          <%    Dim strClassC    %>
          <%    Dim strClassR    %>
          <% 	intCount = 0	 %>
          <%    bolLight = False %>
          
          <% If adoRS.State = adStateOpen Then %>
          <% While (adoRS.EOF = False) %>
          <%    If bolLight = True Then       
                   strClass = "profile_light" 
                   strClassC = "profile_light_ctr" 
                   strClassR = "profile_light_right" 

                Else
                   strClass = "profile_dark"
                   strClassC = "profile_dark_ctr" 
                   strClassR = "profile_dark_right" 
                   
                End If
          %>
          <tr>
            <td class="<%=strClassC %>" style="width: 14%"><font size="2"><%=adoRS.Fields("loc_cd").Value %></font></td>
            <td class="<%=strClassC %>" style="width: 14%"><font size="2"><%=adoRS.Fields("rtrn_loc_cd").Value %></font></td>
            <td class="<%=strClass %>" style="width: 14%"><font size="2"><%=adoRS.Fields("car_type_cd").Value %></font></td>
            <td class="<%=strClass %>" style="width: 14%"><font size="2"><%=adoRS.Fields("car_desc").Value %></font></td>
            <td class="<%=strClassC %>" style="width: 14%"><font size="2"><%=adoRS.Fields("status_cd").Value %></font></td>
            <td class="<%=strClassC %>" style="width: 14%"><font size="2"><%=adoRS.Fields("return_dttm").Value %></font></td>
            <td class="<%=strClassR %>" style="width: 14%"><font size="2"><%=adoRS.Fields("car_count").Value %></font></td>
            <td class="<%=strClassC %>" style="width: 14%"><font size="2"><%=adoRS.Fields("begin_dt").Value %></font></td>
            <td class="<%=strClassC %>" style="width: 14%"><font size="2"><%=adoRS.Fields("end_dt").Value %></font></td>
            <td class="boxtitle" style="width: 14%">&nbsp;</td>
            
  
            <td class="boxtitle" style="width: 14%">&nbsp;</td>
            
  
            <td class="boxtitle" style="width: 14%">&nbsp;</td>
            
  
            <td class="style14" style="height: 11px"><a  href="rezcentral_fleet_adjustment.asp?delete=true&car_id=<%=adoRS.Fields("car_id").Value %>">&nbsp;&nbsp;delete</a></td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%    bolLight = Not bolLight %>
          <%   adoRS.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          <tr>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="style13" style="height: 11px">
			&nbsp;</td>
            <td class="style13" style="height: 11px">
			&nbsp;</td>
            <td class="style13" style="height: 11px">
			&nbsp;</td>
            <td class="style13" style="height: 11px">
			&nbsp;</td>
            <td class="style12" style="height: 11px">&nbsp;</td>
          </tr>
          <tr>
            <td class="boxtitle" style="width: 14%">
		
			<select size="1" name="loc_cd" style="width: 51px">
                   <%   While (adoRS1.EOF = False) %>
 		                   
		                    <option ><%=adoRS1.Fields("city_cd").Value %></option>
		 		   <%     adoRS1.MoveNext %>
		           <%   Wend %>
		   </select>			
			
			</td>
            <td class="boxtitle" style="width: 14%">
			
			
			<select size="1" name="rtrn_loc_cd" style="width: 51px">
                   <%   While (adoRS2.EOF = False) %>
 		                   
		                    <option ><%=adoRS2.Fields("city_cd").Value %></option>
		 		   <%     adoRS2.MoveNext %>
		           <%   Wend %>
		   </select>			
			
			
			</td>
            <td class="boxtitle" style="width: 14%">
			<select size="1" name="car_type_cd" style="width:75;" >
		           <% While (adoRS3.EOF = False)  %>
		                  <option ><%=adoRS3.Fields("car_type_cd").Value %></option>
		 		   <%   adoRS3.MoveNext %>
		           <% Wend %>
					

            </select>
			</td>
            <td class="boxtitle" style="width: 14%">
			<input name="car_desc" type="text" size="20" value=""/></td>
            <td class="boxtitle" style="width: 14%">
			<input name="status_cd" type="text" size="1" value="0"/></td>
            <td class="boxtitle" style="width: 14%">
			<input name="return_dttm" type="text" size="10" value="mm/dd/yyyy" style="width: 80px"/></td>
            <td class="boxtitle" style="height: 11px">
			<input name="car_count"  type="text" size="10" value="1" style="width: 80px" align="right" /></td>
            <td class="boxtitle" style="height: 11px">
			<input name="begin_dt" type="text" size="10" value="mm/dd/yyyy" style="width: 80px"/></td>
            <td class="boxtitle" style="height: 11px">
			<input name="end_dt" type="text" size="10" value="mm/dd/yyyy" style="width: 80px"></td>
            <td class="boxtitle" style="height: 11px">
			&nbsp;</td>
            <td class="style12" style="height: 11px">&nbsp;</td>
          </tr>
          <tr>
            <td class="boxtitle" style="width: 14%">
		
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			
			
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="width: 14%">
			&nbsp;</td>
            <td class="boxtitle" style="height: 11px">
			&nbsp;</td>
            <td class="boxtitle" style="height: 11px">
			&nbsp;</td>
            <td class="boxtitle" style="height: 11px">
			&nbsp;</td>
            <td class="boxtitle" style="height: 11px">
			&nbsp;</td>
            <td class="style12" style="height: 11px">&nbsp;</td>
          </tr>
          </table>
		<input name="Add" type="submit" value="Add Adjustment"><br/>
		<br/>
		Total: <%=intCount %>
        </div>
        <br/>
<br/>
</form>
<form name="update_utilization" action="" method="get">
	<div class="style11">The button below will invoke a full utilization 
		recalculation<br/>for all locations and car types. <br/>
<input name="Add" type="submit" value="Update Utilization" onclick="skm_LockScreen('We are processing your request...');" /><br/>
This process may take up to 15 minutes to complete<br/>
	</div>
	<div id="skm_LockPane" class="LockOff"></div> 
</form>
  <table border="0" style="width: 800px;" bordercolor="#FFFFFF" height="4" id="table1" align="center" class="style7">
    <tr>
      <td background="images/ruler.gif"></td>
      
    </tr>
  </table>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp; </font>

<p align="center">&nbsp;</p>
<p class="style12">
u: <%=strUserId %></p><br/>
</p>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
<% Set adoRS = Nothing
   Set adoCmd = Nothing 
%>