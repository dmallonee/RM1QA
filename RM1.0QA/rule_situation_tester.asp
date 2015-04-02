<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" -->
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 180

	On Error Resume Next

	Dim strConn	
	Dim adoCmd	
	Dim adoRS

	Dim adoPrices
	Dim strUserId
	Dim intRuleId
	Dim strAlertDesc
	Dim datBeginDate
	Dim blnSuccess
	Dim blnIsDollar
	
	blnIsDollar = True

    'On Error Resume Next

	'intRuleId = Request.Form("rateruleid")	
	
	strConn = Session("pro_con")
	
	If Request.Form("situation_cd") > 0 Then
	
	    intSituationCd = Request.Form("situation_cd")
	    blnIsDollar = Request.Form("is_dollar")
    	
	    Rem Get the data sources
	    Set adoCmd = CreateObject("ADODB.Command")

	    adoCmd.ActiveConnection =  strConn
	    adoCmd.CommandText = "car_rate_rule_tester"
	    adoCmd.CommandType = 4


	    If IsNumeric(Request.Form("rt_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@rt_amt",        6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@rt_amt",        6, 1, 0, Request.Form("rt_amt"))
		End If

	    If IsNumeric(Request.Form("situation_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_amt", 6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_amt", 6, 1, 0, Request.Form("situation_amt"))
		End If

	    If IsNumeric(Request.Form("max_rt_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@max_rt_amt",    6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@max_rt_amt",    6, 1, 0, Request.Form("max_rt_amt"))
		End If

	    If IsNumeric(Request.Form("min_rt_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@min_rt_amt",    6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@min_rt_amt",    6, 1, 0, Request.Form("min_rt_amt"))
		End If

	    If IsNumeric(Request.Form("add_rt_amt")) = False Then
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@add_rt_amt",    6, 1, 0, 0)
	    Else
		    adoCmd.Parameters.Append adoCmd.CreateParameter("@add_rt_amt",    6, 1, 0, Request.Form("add_rt_amt"))
		End If

	    adoCmd.Parameters.Append adoCmd.CreateParameter("@is_dollar",    11, 1, 0, Request.Form("is_dollar"))
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@situation_cd",  3, 1, 0, Request.Form("situation_cd"))
	    adoCmd.Parameters.Append adoCmd.CreateParameter("@success",      11, 3, 0, Null)


	    adoCmd.Execute

	    blnSuccess = adoCmd.Parameters("@success").Value

        Set adoCmd = Nothing

        If err.number <> 0 Then
	        pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	        response.write "<b>VBScript Errors Occured!<br>"
	        response.write "</b><br>"
	        response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	        response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	        response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	        response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	        response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
        End If


    End If

%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Rate rule situation tester</title>
<style type="text/css">
.style1 {
	font-size: x-small;
}
</style>
</head>

<body bgcolor="#D8DEE1" style="font-family: Verdana; font-size: 10pt">

<p>&nbsp;</p>
    <h2>
        Rate Rule Situation Tester</h2>
      <form method="POST" name="situation_tester" action="rule_situation_tester.asp">

  <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="717" id="table1" >
    <tr>
      <td width="160" height="22">Competitor:</td>
      <td width="557" height="22" colspan="2">
      &nbsp;</td>
    </tr>
    <tr>
      <td width="160" height="22"><font size="2">&nbsp; Max. rate:</font></td>
      <td width="557" height="22" colspan="2">
        <p>
        <input type="text" name="max_rt_amt" size="20" style="text-align: right; width:75;"  value="<%=Request.Form("max_rt_amt") %>" ></p></td>
    </tr>
    <tr>
      <td width="160" height="22"><font size="2">&nbsp; Min. rate:</font></td>
      <td width="557" height="22" colspan="2">
      <input type="text" name="min_rt_amt" size="20"  style="text-align: right; width:75;"  value="<%=Request.Form("min_rt_amt") %>" id="Text1" onclick="return Text1_onclick()" ></td>
    </tr>
    <tr>
      <td width="160" height="22"><font size="2">&nbsp; Extra rate:</font></td>
      <td width="557" height="22" colspan="2">
      <input type="text" name="add_rt_amt" size="20"  style="text-align: right; width:75;"  value="<%=Request.Form("add_rt_amt") %>" >
		<span class="style1">(this rate is only used when the two lowest 
		competitors are required)</span></td>
    </tr>
    <tr>
      <td width="160" height="22">&nbsp;</td>
      <td width="557" height="22" colspan="2">
      <span
          style="font-size: 2px"></span></td>
    </tr>
    <tr>
      <td width="160" height="22"><font size="2">Comparison rate:</font></td>
      <td width="557" height="22" colspan="2">
      <input type="text" name="rt_amt" size="20"  style="text-align: right; width:75;"  value="<%=Request.Form("rt_amt") %>" >
		<span class="style1">(usually this is your rate)</span></td>
    </tr>
    <tr>
      <td width="160" height="22">&nbsp;</td>
      <td width="557" height="22" colspan="2">
      &nbsp;</td>
    </tr>
    <tr>
<font color="#080000">
      <td width="160" height="22">
<font size="2">Situation:</font></td>
      <td width="557" height="22" colspan="2">
      <font size="2" color="#080000">
		<select size="1" name="situation_cd" style="width:370; font-family:Verdana; font-size:10pt; height:24">
      
      <% If intSituationCd = 0 Then	%>
	      <option selected value="0">(None selected)</option>
	  <% Else %>
	      <option value="0">(None selected)</option>
	  <% End If %>
 
      <% If intSituationCd = 1 Then	%>
	      <option selected value="1">NONE - Set rate to the response amount</option>
	  <% Else %>
	      <option value="1">NONE - Set rate to the response amount</option>
	  <% End If %>
	  <!-- 	  
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
	  -->
      <% If intSituationCd = 4 Then	%>
	      <option selected value="4">If rate is not more than (all competitors) by at least</option>
	  <% Else %>
	      <option value="4">If rate is not more than (all competitors) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 5 Then	%>
	      <option selected value="5">If rate is not more than (any competitors) by at least</option>
	  <% Else %>
	      <option value="5">If rate is not more than (any competitors) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 6 Then	%>
	      <option selected value="6">If rate is not more than (custom) by at least</option>
	  <% Else %>
	      <option value="6">If rate is not more than (custom) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 7 Then	%>
	      <option selected value="7">If rate is equal to (all competitors)</option>
	  <% Else %>
	      <option value="7">If rate is equal to (all competitors)</option>
	  <% End If %>

      <% If intSituationCd = 8 Then	%>
	      <option selected value="8">If rate is equal to (any competitor)</option>
	  <% Else %>
	      <option value="8">If rate is equal to (any competitor)</option>
	  <% End If %>

      <% If intSituationCd = 9 Then	%>
	      <option selected value="9">If rate is equal to (custom)</option>
	  <% Else %>
	      <option value="9">If rate is equal to (custom)</option>
	  <% End If %>
	  <!--
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
	  -->
      <% If intSituationCd = 14 Then	%>
	      <option selected value="14">If rate is not less than (all competitors) by at least</option>
	  <% Else %>
	      <option value="14">If rate is not less than (all competitors) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 16 Then	%>
	      <option selected value="16">If rate is not less than (any competitor) by at least</option>
	  <% Else %>
	      <option value="16">If rate is not less than (any competitor) by at least</option>
	  <% End If %>

      <% If intSituationCd = 15 Then	%>
	      <option selected value="15">If rate is not less than (custom) by at least</option>
	  <% Else %>
	      <option value="15">If rate is not less than (custom) by at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 17 Then	%>
	      <option selected value="17">If rate is not equal to (any competitor)</option>
	  <% Else %>
	      <option value="17">If rate is not equal to (any competitor)</option>
	  <% End If %>
	  
      <% If intSituationCd = 18 Then	%>
	      <option selected value="18">If rate is not equal to (all competitors)</option>
	  <% Else %>
	      <option value="18">If rate is not equal to (all competitors)</option>
	  <% End If %>
	  
      <% If intSituationCd = 19 Then	%>
	      <option selected value="19">If rate is not equal to (custom)</option>
	  <% Else %>
	      <option value="19">If rate is not equal to (custom)</option>
	  <% End If %>
	  <!--
      <% If intSituationCd = 20 Then	%>
	      <option selected value="20">If rate is not equal to  (all competitors) + diff</option>
	  <% Else %>
	      <option value="20">If rate is not equal to  (all competitors) + diff</option>
	  <% End If %>
 	  -->
      <% If intSituationCd = 30 Then	%>
	      <option selected value="30">If the diff. between (all competitors) is at least</option>
	  <% Else %>
	      <option value="30">If the diff. between (all competitors) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 31 Then	%>
	      <option selected value="31">If the diff. between (all competitors) is less than</option>
	  <% Else %>
	      <option value="31">If the diff. between (all competitors) is less than</option>
	  <% End If %>

      <% If intSituationCd = 32 Then	%>
	      <option selected value="32">If the diff. between (any competitor) is at least</option>
	  <% Else %>
	      <option value="32">If the diff. between (any competitor) is at least</option>
	  <% End If %>
	  
      <% If intSituationCd = 33 Then	%>
	      <option selected value="33">If the diff. between (any competitor) is less than</option>
	  <% Else %>
	      <option value="33">If the diff. between (any competitor) is less than</option>
	  <% End If %>

      <% If intSituationCd = 34 Then	%>
	      <option selected value="34">If rate is not less than (all competitors) by exactly</option>
	  <% Else %>
	      <option value="34">If rate is not less than (all competitors) by exactly</option>
	  <% End If %>
	  
      <% If intSituationCd = 35 Then	%>
	      <option selected value="35">If rate is not less than (any competitor) by exactly</option>
	  <% Else %>
	      <option value="35">If rate is not less than (any competitor) by exactly</option>
	  <% End If %>

      <% If intSituationCd = 40 Then	%>
	      <option selected value="40">If (any competitor) rate is less than</option>
	  <% Else %>
	      <option value="40">If (any competitor) rate is less than</option>
	  <% End If %>

      <% If intSituationCd = 41 Then	%>
	      <option selected value="41">If (all competitor) rates are less than</option>
	  <% Else %>
	      <option value="41">If (all competitor) rate are less than</option>
	  <% End If %>

      <% If intSituationCd = 42 Then	%>
	      <option selected value="42">If (any competitor) rate is equal to</option>
	  <% Else %>
	      <option value="42">If (any competitor) rate is equal to</option>
	  <% End If %>

      <% If intSituationCd = 43 Then	%>
	      <option selected value="43">If (all competitor) rates are equal to</option>
	  <% Else %>
	      <option value="43">If (all competitor) rate are equal to</option>
	  <% End If %>

      <% If intSituationCd = 44 Then	%>
	      <option selected value="44">If (any competitor) rate is greater than</option>
	  <% Else %>
	      <option value="44">If (any competitor) rate is greater than</option>
	  <% End If %>

      <% If intSituationCd = 45 Then	%>
	      <option selected value="45">If (all competitor) rates are greater than</option>
	  <% Else %>
	      <option value="45">If (all competitor) rate are greater than</option>
	  <% End If %>

      <% If intSituationCd = 46 Then	%>
	      <option selected value="46">If gap between two lowest competitors is greater than</option>
	  <% Else %>
	      <option value="46">If gap between two lowest competitors is greater than</option>
	  <% End If %>

      <% If intSituationCd = 47 Then	%>
	      <option selected value="47">Is comparison rate closed?</option>
	  <% Else %>
	      <option value="47">Is comparison rate closed?</option>
	  <% End If %>

      <% If intSituationCd = 48 Then	%>
	      <option selected value="48">Is comp set rate closed?</option>
	  <% Else %>
	      <option value="48">Is comp set rate closed?</option>
	  <% End If %>

      </select></font></td>
</font>
    </tr>
    <tr>
<font color="#080000">
      <td width="160" height="23">
<font color="#080000">
      <font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; 
      Amount:</font></font></td>
      <td width="89" height="23">
      <input type="text" name="situation_amt" size="20" style="text-align: right; width:75; font-family:Verdana; font-size:10pt; text-align:right; height:22" value="<%=Request.Form("situation_amt") %>"></td>
      
      
</font>
      <td width="468" height="23" align="left" style="text-align: left">
<font color="#080000">
      <% If blnIsDollar Then %>
      <input type="radio" value="11" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1" checked="CHECKED"><font face="Verdana" size="2"><label for="is_dollar1">Dollar </label>
      <input type="radio" value="01" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar2"  ><label for="is_dollar2">Percentage</label><br>
      <% Else %>
      <input type="radio" value="12" name="is_dollar" style="font-family:Verdana; font-size:10pt" id="is_dollar1"><label for="is_dollar1">Dollar </label>
      <input type="radio" value="02" name="is_dollar" style="font-family:Verdana; font-size:10pt"  id="is_dollar2"><label for="is_dollar2">Percentage</label><br>
      <% End If %></font></font></td>
      
      
    </tr>
    <tr>
      <td width="160" height="23">
&nbsp;</td>
      <td width="557" height="23" colspan="2">
      &nbsp;</td>
      
      
    </tr>
    <tr>
      <td width="160" height="23">
&nbsp;</td>
      <td width="557" height="23" colspan="2">
      <input type="submit" value="Evalute Situation" name="submit"></td>
      
      
    </tr>
  </table>
  </form> 
    <br />
    <% If Request.Form("situation_cd") > 0 Then %>
    
    <% If IsNull(blnSuccess) Then %>
    <span style="font-size: 11pt;"><strong>ERROR</strong></span> --
    This situation is not functioning properly, please do not use it.<br>
    
    <% ElseIf blnSuccess Then %>
    <span style="font-size: 11pt; color: #009900"><strong>TRUE</strong></span> --
    The situation is true, the response will be performed, and if there is a follow-on rule for the true
    outcome of the situation, it will be performed.
    <% Else %>
    <br>
    <span style="font-size: 11pt; color: #cc0000"><strong>FALSE</strong></span> --
    The situation is false, the response will not be performed, and if there is a follow-on
    rule for
    the false outcome of the situation, it will be performed.<br />
    <% End If %>
    
    <% End If %>
<br>
<font size="1">
rule id: <%=Request.Form("situation_cd") %><br>
</font>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>