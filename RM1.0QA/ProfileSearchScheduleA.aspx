<%@ Page Language="VB" MasterPageFile="~/Admin.master" AutoEventWireup="false" CodeFile="ProfileSearchScheduleA.aspx.vb" Inherits="ProfileSearchScheduleA" title="Untitled Page" EnableEventValidation="false" ValidateRequest="false" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cphMain" Runat="Server">
<asp:Label ID="lblMessage" runat="Server"></asp:Label>
<p>
&nbsp;[create or edit] [<a target="_self" href="ProfileSearchScheduleAll.aspx">view all schedules</a>]&nbsp;
<p align="center"><font size="5" color="#384F5B">Create or Edit Search Schedules</font></p>
<p align="center"><font size="2" color="#384F5B">Edit an existing schedule by 
selecting one from the list below</font></p>
 
	<div align="center">
 
	<table border="0" width="640" id="table1" bgcolor="#CFD7DB">
		<tr>
			<td width="128"><font size="2">Edit Schedule: </font></td>
			<td>&nbsp;</td>
			<td width="493" style="text-align: left">
			<asp:DropDownList ID="ddlSchedule" runat="server" Width="400"></asp:DropDownList>
			
			</td>
		</tr>
		<tr>
			<td width="128">&nbsp;</td>
			<td>&nbsp;</td>
			<td width="493">&nbsp;</td>
		</tr>
		<tr>
			<td width="128">&nbsp;</td>
			<td>&nbsp;</td>
			<td width="493">
			<asp:Button ID="btnEdit" runat="Server" Text="  Next &gt;&gt;  "/>
			 </td>
		</tr>
	</table>
	</div>
  <p align="center">&nbsp;</p>

 
	<div align="center">
 
	<table border="0" width="640" id="new_schedule" bgcolor="#CFD7DB">
		<tr>
			<td  width="128"><font size="2">New Schedule:</font></td>
			<td>&nbsp;</td>
			<td width="493" style="text-align: left">
			<asp:TextBox ID="txtNewName" runat="Server" Width="400"></asp:TextBox>
						
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td>
			<asp:Button ID="btnNew" runat="Server" Text="  Next &gt;&gt;  " />
			 </td>
		</tr>
	</table>
	</div>
  <p align="center">&nbsp;</p>
	

<p align="center">&nbsp;</p>
<div align="center">
	<table border="0" width="640" id="table3">
		<tr>
			<td style="text-align: left"><font size="2"><b>Directions:</b> Either select an existing 
			schedule from the list above or select &quot;New Schedule&quot; from the 
			drop-down list to create a new search schedule for the profile that 
			you are currently viewing. If you decide to create a new 
			schedule you will be able to enter a descriptive name for it in this 
			or the&nbsp; 
			next step.</font></td>
		</tr>
		<tr>
			<td>&nbsp;</td>
		</tr>
	</table>
</div>
<p align="center">&nbsp;</p>
<p align="center">&nbsp;</p>
</asp:Content>

