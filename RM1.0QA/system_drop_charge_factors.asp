<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- #INCLUDE FILE="inc/login_check.asp" -->
<% 	Response.Expires = -1  
   	Response.cachecontrol="private" 
   	Response.AddHeader "pragma", "no-cache" 
   
   	On error resume next

   	Server.ScriptTimeout = 180

	Dim strSelected 

    strUserId     = Request.Cookies("rate-monitor.com")("user_id")
	strOrigCityCd = Request.Form("orig_city_cd")
	strDestCityCd = Request.Form("dest_city_cd")
	strCarTypeCd  = Request.Form("car_type_cd")

	curFeeAmt = 0
	curTaxAmt = 0

	strConn = Session("pro_con")
	
	If Request.Form("update") = "true" Then
	
		Set adoCmd = CreateObject("ADODB.Command")

		adoCmd.ActiveConnection =  strConn
		adoCmd.CommandText = "car_drop_charge_variable_insert"
		adoCmd.CommandType = adCmdStoredProc
	
		adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",        3, 1, 0, strUserId)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@orig_city_cd", 200, 1, 6, strOrigCityCd)
		'adoCmd.Parameters.Append adoCmd.CreateParameter("@dest_city_cd", 200, 1, 6, strDestCityCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd",  200, 1, 4, strCarTypeCd)
		adoCmd.Parameters.Append adoCmd.CreateParameter("@fee_amt",        6, 1, 0, Request.Form("fee_amt"))
		adoCmd.Parameters.Append adoCmd.CreateParameter("@tax_amt",        6, 1, 0, Request.Form("tax_amt")/100)
	
		adoCmd.Execute	

		
		If err.number <> 0 Then
		   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
		   response.write "<b>Insert Error Occured<br>"
		   response.write parm_msg & "</b><br>"
		   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
		   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
		   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
		   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
		   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
		   response.write pad & "Data= <b>" & strUserId & strCityCd & strCarTypeCd & Request.Form("fee_amt") & Request.Form("tax_amt") & "</b><br><hr>"
	
		End If

	
	End If
	
	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "user_city_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",  3, 1,  0, strUserId)
		
	Set adoRSorig = adoCmd.Execute
	Set adoRSdest = adoCmd.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>City Error Occured<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If

	adoCmd.CommandText = "car_type_select"
	'We dont need to set the user id because it is still set from above.
		
	Set adoRS1 = adoCmd.Execute

	If strOrigCityCd = "" Then
		strOrigCityCd = adoRSorig.Fields("city_cd").Value
	End If


	If strDestCityCd = "" Then
		strDestCityCd = adoRSdest.Fields("city_cd").Value
	End If

	If strCarTypeCd = "" Then
		strCarTypeCd = adoRS1.Fields("car_type_cd").Value
	End If

	Rem Get the data sources
	Set adoCmd = CreateObject("ADODB.Command")

	adoCmd.ActiveConnection =  strConn
	adoCmd.CommandText = "car_drop_charge_variable_select"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",        3, 1, 0, strUserId)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@orig_city_cd", 200, 1, 6, strOrigCityCd)
	'adoCmd.Parameters.Append adoCmd.CreateParameter("@dest_city_cd", 200, 1, 6, strDestCityCd)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@car_type_cd",  200, 1, 4, strCarTypeCd)
		
	Set adoRS2 = adoCmd.Execute

	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>Drop Charge Select Error Occured<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If



	If adoRS2.EOF = True Then
		curFeeAmt = 0
		curTaxAmt = 0
	Else
		curFeeAmt = adoRS2.Fields("fee_amt").Value
		curTaxAmt = adoRS2.Fields("tax_amt").Value
		
	End If	
		
	If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write parm_msg & "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If
  
%>    
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. |&nbsp; One Way Settings</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<script type="text/javascript" language="JavaScript" src="inc/sitewide.js" ></script>
<script type="text/javascript" language="JavaScript" src="inc/multiple_select_support2.js"></script>
<script type="text/javascript" language="JavaScript" >
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

function update_car_type_list() { 

	var list = new String("")
	
	for (i = 0; i < document.car_group_detail.selected_car_types.options.length;i++)
	{
		if (i == 0)
		{
	      list = document.car_group_detail.selected_car_types.options[i].value
		}
		else
		{
	      list = list + ',' + document.car_group_detail.selected_car_types.options[i].value
		}
	}

    document.car_group_detail.car_type_list.value = list 
    //document.search_criteria.selected_companies.options.text; 
     
} 


function removeText(what){
	var what=document.getElementById(what);
	whatChild=what.removeChild(what.childNodes[0]);
}

function replaceText(what,hlaska){
	removeText(what);
	var newText=document.createTextNode(hlaska);
	document.getElementById(what).appendChild(newText);
}

function update_city_cd() {
	replaceText('city_cd', document.getElementById('grp_city_cd').options[document.getElementById('grp_city_cd').selectedIndex].value);
}

function update_grp_car_type() {
	replaceText('car_type_cd', document.getElementById('grp_car_class').options[document.getElementById('grp_car_class').selectedIndex].value);
}

//-->
</script>

<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<style type="text/css"  >
<!--
.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#B2BEC4; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.style2 {
	border-collapse: collapse;
}
.style3 {
	border-style: solid;
	border-width: 0;
}
.style5 {
	text-align: right;
}
.style6 {
	text-align: right;
	font-size: x-small;
}
.style8 {
	border: 1px solid #3C5967;
	border-collapse: collapse;
}
.style9 {
	font-size: x-small;
}
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
<!-- #INCLUDE FILE="inc/page_header_buttons.htm" -->
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
city codes]</a>&nbsp;<a href="javascript:not_enabled()">[system status]</a>
<a href="system_proxy.asp">[proxy 
management]</a> <a href="system_utilization.asp">[utilization settings]</a><b> </b>
<a href="system_utilization_car_groups.asp">[utilization car groups]</a><b> 
[one-way settings]</b></font><br>
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
<p>&nbsp;
<FORM METHOD="post" NAME="city_select"   > 
<table border="0" cellpadding="0" bordercolor="#111111" width="500" align="center" class="style2">
          <tr>
           <td width="100%" class="style3" colspan="3"><font size="2"><b>
           Directions:</b> To manage the fees and tax amounts for one-way 
			charges please 
           use this page. Select a city code and a car class, then click the view button to 
           display the selected city and class's one-way drop charge amounts. To save 
			new values or change current values, enter the new values and click 
			the update button. The calculation used is as follows:<br>
			<br>
			<strong>Drop charge = [ (Total Price – Fee) / (1 + Tax) ] – base rate</strong> <br></font>
		   <br>
		   </tr>
		   <tr>
				<td><font size="2">Origin City:</font></td>
				<td>
					<select size="1" name="orig_city_cd" id="orig_city_cd" style="width:75;" tabindex="1" >
		           <% While (adoRSorig.EOF = False) 
		                If adoRSorig.Fields("city_cd").Value = strOrigCityCd Then %>
		                  <option selected ><%=adoRSorig.Fields("city_cd").Value %></option>		           
		           <%   Else %>	 
		                  <option ><%=adoRSorig.Fields("city_cd").Value %></option>
		           <%   End If %>
		 		   <%   adoRSorig.MoveNext %>
		           <% Wend %>
		           </select>
		        </td>
				<td rowspan="5">
					<font size="2"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font> 
		            <input type="submit" value=" View " name="view" class="rh_button" tabindex="3">
		        </td>
			</tr>
		   <tr>
				<td>&nbsp;</td>
				<td>
					&nbsp;</td>
			</tr>
		<!--	
		   <tr>
				<td><font size="2">Dest. City:</font></td>
				<td>
					<select size="1" name="city_cd0" id="grp_city_cd0" style="width:75;" tabindex="1" onchange="update_city_cd()">
		           <% While (adoRSdest.EOF = False) 
		                If adoRSdest.Fields("city_cd").Value = strCityCd Then %>
		                  <option selected ><%=adoRSdest.Fields("city_cd").Value %></option>		           
		           <%   Else %>	 
		                  <option ><%=adoRSdest.Fields("city_cd").Value %></option>
		           <%   End If %>
		 		   <%   adoRSdest.MoveNext %>
		           <% Wend %>
		           </select>
		        </td>
			</tr>

		   <tr>
				<td colspan="2" class="style3" style="height: 5px">&nbsp;</td>
				<td class="style3" style="height: 5px">
					&nbsp;</td>
			</tr>
		 -->
			<tr>
				<td><font size="2">Car Class:</font></td>
				<td>
					<select size="1" name="car_type_cd" id="grp_car_class" style="width:75;" tabindex="2" onchange="update_grp_car_type();">
		           <% While (adoRS1.EOF = False) 
		                If adoRS1.Fields("car_type_cd").Value = strCarTypeCd Then %>
		                  <option selected ><%=adoRS1.Fields("car_type_cd").Value %></option>		           
		           <%   Else %>	 
		                  <option ><%=adoRS1.Fields("car_type_cd").Value %></option>
		           <%   End If %>
		 		   <%   adoRS1.MoveNext %>
		           <% Wend %>
					

                   </select>
                </td>
			</tr>
		   	</table>
		<input type="hidden" name="update" value="false">
		</form>

           <p>&nbsp;</p>

<FORM METHOD="post" NAME="detail" action=""   > 

<table cellpadding="0" width="500" align="center" class="style8">
                      	
                  <tr>
                     <td style="width: 149px" class="style5">&nbsp;</td>
                    <td >&nbsp;</td>                    
                    <td >&nbsp;</td>
                  </tr>
                  <tr>
                     <td style="width: 149px" class="style6">Fee amount ($):&nbsp;&nbsp; </td>
                    <td >
					<input name="fee_amt" type="text" value='<%=FormatNumber(curFeeAmt, 2) %>' style="text-align: right; width: 83px;"></td>                    
                    <td >&nbsp;</td>
                  </tr>
                  <tr>
                    <td style="width: 149px" class="style6">Tax amount (%):&nbsp;&nbsp;</td>
                    <td >
					<input name="tax_amt" type="text" value='<%=FormatNumber(curTaxAmt * 100, 2) %>' style="text-align: right; width: 83px;"></td>                    
                    <td >&nbsp;</td>
                  </tr>
                  <tr>
                    <td style="width: 149px" class="style6">Car Class:&nbsp;&nbsp;</td>
                    <td class="style9">
					<input type="text" name="car_type_cd" id="car_type_cd" value="<%=strCarTypeCd %>" style="width: 83px; background-color:silver" readonly="readonly" >( 
					to change - select new values above</td>                    
                    <td >&nbsp;</td>
                  </tr>
                  <tr>
                    <td style="width: 149px" class="style6">Orig. City Code:&nbsp;&nbsp;</td>
                    <td class="style9">
					<input type="text" name="orig_city_cd" id="orig_city_cd" value="<%=strOrigCityCd %>" style="width: 83px; background-color:silver" readonly="readonly" > 
					and click the view button )</td>                    
                    <td >&nbsp;</td>
                  </tr>
                  <!-- 
                  <tr>
                    <td style="width: 149px" class="style6">Dest. City Code:&nbsp;&nbsp;</td>
                    <td >
					<input type="text" name="dest_city_cd" id="dest_city_cd" value="<%=strDestCityCd %>" style="width: 83px; background-color:silver" readonly="readonly" ></td>                    
                    <td >&nbsp;</td>
                  </tr>
                  -->
                  <tr>
                    <td style="width: 149px" class="style5">&nbsp;</td>
                    <td >
					&nbsp;</td>                    
                    <td >&nbsp;</td>
                  </tr>
                  <tr>
                    <td style="width: 149px">&nbsp;</td>
                    <td ><input type=submit value='Update' name=submit caption="Add To Database" class="rh_button" >&nbsp;&nbsp;</td>                    
                    <td >&nbsp;</td>
                  </tr>
         </table>
        		<br>
			    <br>

        

        
        		<input type="hidden" name="update" value="true">
        		
  			    

        

        
        </FORM>
      

        <p>&nbsp;<p align="center">&nbsp;</p>

  <table border="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="1114" cellspacing="0" height="4" id="table1">
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
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>