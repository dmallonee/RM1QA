<%@ Page Title="" Language="VB" MasterPageFile="~/popup.master" AutoEventWireup="false"
    CodeFile="threshold_maxmin_schedule_a.aspx.vb" Inherits="threshold_maxmin_schedule_a" %>

<asp:Content ID="C1" ContentPlaceHolderID="cphMain" Runat="Server">
    <p align="center">
        &nbsp;</p>
    <p align="center">
        <font size="5" color="#384F5B">Threshold Max. / Min. Schedule</font></p>
    <p align="center">
        <font size="2" color="#384F5B">Edit an existing schedule by selecting one from 
        the list below</font></p>
    <div align="center">
     
        <table border="0" width="640" bgcolor="#CFD7DB">
            <tr>
                <td width="128">
                    <font size="2">Edit/Delete Schedule: </font>
                </td>
                <td>
                    &nbsp;
                </td>
                <td width="493" align="left">
                    <asp:DropDownList ID="ddlSchedules" runat="server" Width="400" 
                        DataTextField="car_rate_rule_schedule_desc" 
                        DataValueField="car_rate_rule_schedule_id">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="128">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td width="493">
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td width="128">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td width="493" align="left">
                    <asp:Button ID="btnEdit" runat="server" Text="  Edit  " Width="80px" /> &nbsp;&nbsp;  
                    <asp:Button ID="btnDelete" runat="server" Text="  Delete  " Width="80px" />

               

                </td>
            </tr>
        </table>
      
    </div>
    <p align="center">
        &nbsp;</p>
    <p align="center">
        <font color="#384F5B">Or create a new schedule by using the option below</font></p>
    <div align="center">
        <table border="0" width="640" bgcolor="#CFD7DB">
            <tr>
                <td width="128">
                    <font size="2">Create Schedule:</font>
                </td>
                <td>
                    &nbsp;
                </td>
                <td width="493" align="left">
                    <asp:TextBox ID="txtNewName" runat="server" Width="400" MaxLength="40"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td>
                    <font size="2">Schedule Type:</font>
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="left">
                    <asp:DropDownList ID="ddlScheduleTypes" runat="server" Width="100">
                        <asp:ListItem Selected="True" Text="Max./Min." Value="6"></asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td align="left">
                    <asp:Button ID="btnNew" runat="Server" Text="  Create  " Width="80px" />
                </td>
            </tr>
        </table>
    </div>
    <p align="center">
        &nbsp;</p>
    <p align="center">
        &nbsp;</p>
    <div align="center">
        <table border="0" width="640">
            <tr>
                <td>
                    <font size="2"><b>Directions:</b> Either select an existing schedule from the 
                    list above or select &quot;New Schedule&quot; from the drop-down list to create a new rule 
                    schedule. If you select a new schedule you will be able to create a descriptive 
                    name for it in the next step</font>
                </td>
            </tr>
            <tr>
                <td>
                    &nbsp;
                </td>
            </tr>
        </table>
    </div>
    <p align="center">
        &nbsp;</p>
    <p align="center">
        &nbsp;</p>
</asp:Content>

