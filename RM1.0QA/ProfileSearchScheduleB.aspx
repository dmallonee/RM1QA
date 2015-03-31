<%@ Page Language="VB" MasterPageFile="~/Admin.master" AutoEventWireup="false" CodeFile="ProfileSearchScheduleB.aspx.vb" Inherits="ProfileSearchScheduleB" title="Untitled Page" ValidateRequest="false" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cphMain" Runat="Server">
<asp:Label ID="lblMessage" runat="server"></asp:Label>
<br />
&nbsp;[<a target="_self" href="ProfileSearchScheduleA.aspx">create or edit</a>] 
<strong>[<a target="_self" href="ProfileSearchScheduleAll.aspx">view 
all schedules</a>]</strong>&nbsp;

<asp:ValidationSummary ID="vsMain" runat="server" />
<div align="center">
<font size="5" color="#384F5B">Search Schedule</font>
</div>
	<table border="0" width="640" id="new_profile" align="center" cellspacing="0" cellpadding="0">
		<tr>
			<td width="113"><font size="2">&nbsp;Schedule Name: </font></td>
			<td>&nbsp;</td>
			<td>
			<asp:TextBox ID="txtName" runat="server" Width="250"></asp:TextBox>
			<asp:RequiredFieldValidator ID="rfvName" runat="server" ControlToValidate="txtName" ErrorMessage="You must enter a name for this schedule." Text="*"></asp:RequiredFieldValidator>
			</td>
		</tr>
		<tr>
			<td width="113">&nbsp;</td>
			<td>&nbsp;</td>
			<td><asp:CheckBox ID="chkSaveCopy" Checked="True" runat="server"  Text="Save as a copy" /></td>
		</tr>
		<tr>
			<td width="113">&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>

		</tr>
		<tr>
		<td class="style2">&nbsp;</td>
		<td width="162" class="style2">
		&nbsp;</td>
                    <td width="365" class="style2">
                  
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>
		<tr>
		<td class="style2">&nbsp; Time:</td>
		<td width="162" class="style2">
		<asp:RadioButton ID="rdoFixed" runat="Server" Text="Fixed:" GroupName="rdoTime" />
		</td>
                    <td width="365" class="style2">
                  <asp:DropDownList ID="ddlScheduledTime" runat="Server"></asp:DropDownList>
                  
                    </td>
                    <td class="style2"></td>
                  </tr>
                  <tr>
                  	<td class="style2"></td>
                  	<td class="style2">&nbsp;&nbsp;&nbsp;&nbsp;or</td>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2"></td>
                  </tr>
                  <tr>
                  	<td class="style2"></td>
                  	<td class="style2">
                  	<asp:RadioButton ID="rdoRandom" runat="server" GroupName="rdoTime" Text="Random:" />
                  	
                  	<td class="style2">Between:
                    <asp:DropDownList ID="ddlScheduledTime0" runat="server"></asp:DropDownList>
                    &nbsp; and&nbsp;
                    <asp:DropDownList ID="ddlScheduledTime1" runat="server"></asp:DropDownList>
                    </td>
                  	<td class="style2"></td>
                  </tr>
                  <tr>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2">&nbsp;</td>
                  	<td class="style2">&nbsp;</td>
                  </tr>
                  <tr>
                  	<td>&nbsp;</td>
                  	<td>&nbsp;</td>
                  	<td>&nbsp;</td>
                  	<td>&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="174" class="style2">
                    &nbsp;
                    Days:</td>
                    <td width="162" class="style2">
                    <asp:RadioButton ID="rdoWeek" GroupName="rdoDays" runat="Server" Text="Weekdays:" />
                    
                    <td width="395">
                    <table width="381" border="0" cellpadding="2" cellspacing="0" >
                      <tr valign="bottom">
                        <td width="20" class="style2">
                        <asp:CheckBox ID="chkDowSunday" runat="server" />
		                
                        </td>
                        <td width="29" class="style2">
                        Sun
                        </td>
                        <td width="20" class="style2">
                        <asp:CheckBox ID="chkDowMonday" runat="server" />
		                
                        </td>
                        <td width="25" class="style2">
                        Mon
                        </td>
                        <td width="20" class="style2">
                        <asp:CheckBox ID="chkDowTuesday" runat="server" />
		                </td>
                        <td width="31" class="style2">
                        Tue </td>
                        <td width="20" class="style2">
                        <asp:CheckBox ID="chkDowWednesday" runat="server" />
		                
                        </td>
                        <td width="25" class="style2">
                        Wed
                        </td>
                        <td width="20" class="style2">
                        
		                <asp:CheckBox ID="chkDowThursday" runat="server" />
                        </td>
                        <td width="19" class="style2">
                        Thu
                        </td>
                        <td width="20" class="style2">
                        
		                <asp:CheckBox ID="chkDowFriday" runat="server" />
                       </td>
                        <td width="20" class="style2">
                        Fri</td>
                        <td width="20" class="style2">
                        
		                <asp:CheckBox ID="chkDowSaturday" runat="server" />
                       </td>
                        <td width="20" class="style2">
                        Sat</td>
                      </tr>
                    </table>
                    </td>
                    <td class="style2"></td>
                  </tr>
		
	              <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    or</td>
                    <td width="395" class="style2">
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    <asp:RadioButton ID="rdoDayFixed" GroupName="rdoDays" Text="Fixed:" runat="Server" />
                    </td>
                    <td width="395" class="style2">
                    <asp:TextBox ID="txtFixedDate" runat="server"></asp:TextBox>
                     mm/dd/yyyy</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    or</td>
                    <td width="395" class="style2">
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    <asp:RadioButton ID="rdoMonthly" runat="server" GroupName="rdoDays" Text="Monthly:" />
                    </td>
                    <td width="395" class="style2">
                    The <asp:DropDownList ID="ddlDayOfMonth" runat="server"></asp:DropDownList> day of the month</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

                  <tr>
                    <td width="174" class="style2">
                    &nbsp;</td>
                    <td width="162" class="style2">
                    &nbsp;</td>
                    <td width="395" class="style2">
                    &nbsp;</td>
                    <td class="style2">&nbsp;</td>
                  </tr>

	</table>
	<p class="style1"><br>
	<asp:HiddenField ID="schedule_type" runat="Server" />
	<asp:HiddenField ID="schedule_id" runat="Server" />
	<asp:Button ID="btnSubmit" runat="Server" Text="Submit" />
	
  <p align="center">&nbsp;</p>
	
<p align="center">&nbsp;</p>
<div align="center">
	<table border="0" width="500" id="table3">
		<tr>
			<td><font size="2">Directions: Enter the rate amount minimums in the 
			min field, and enter the rate amount maximum that you want the rules 
			to observe in the max field. Do this for each car type. These fields 
			are strictly optional and if you decide not to enter values your 
			rules will still process; however, you will not have your rates 
			limited by the minimums and maximums set here.</font>


			<p>&nbsp;</p></td>
		</tr>
		<tr>
			<td><font size="2">Note, if you do not see a car type here that you 
			need in your account, please contact Customer Support at
			<a href="mailto:support@rate-highway.com">support@rate-highway.com</a> 
			and let them know the car type(s) you would like added to your 
			account.</font></td>
		</tr>
	</table>
</div>

</asp:Content>

