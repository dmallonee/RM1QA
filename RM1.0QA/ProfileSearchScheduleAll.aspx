<%@ Page Language="VB" MasterPageFile="~/Admin.master" AutoEventWireup="false" CodeFile="ProfileSearchScheduleAll.aspx.vb" Inherits="ProfileSearchScheduleAll" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cphMain" Runat="Server">
<asp:Label ID="lblMessage" runat="Server"></asp:Label>
<br />
&nbsp;[<a target="_self" href="ProfileSearchScheduleA.aspx">create or edit</a>] 
<strong>[view 
all schedules]</strong>&nbsp;

<div align="center"><font size="5" color="#384F5B">Search Schedules &amp; Groups</font></div>
<div align="center">
        <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="600" id="AutoNumber1" background="images/alt_color.gif">
          <tr>
           <td width="100%" class="boxtitle" style="border-style: solid; border-width: 0; background-color: #FFFFFF;" colspan="3">
			<div style="text-align: left">
				<font size="2"><b>
           Directions:</b> You can use this page to review all the schedules for 
				your account. You may also create new groups of schedules or 
				modify the members of existing groups by checking or un-checking 
				the schedules listed in the table below</font></div>
			<p>
           &nbsp;</p>
           </td>
           
          </tr>
         
          <tr>
           <td style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF; width: 97px;" >
         <font size="2">Choose Group:</td>
           <td  style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF; width: 97px;" >
		<asp:DropDownList ID="ddlScheduleGroup" runat="Server" DataTextField="schedule_grp_desc" DataValueField="schedule_grp_id" AutoPostBack="true" Width="294px"></asp:DropDownList>
			</td>
           <td style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF; width: 97px;" >
		&nbsp;</td>
          </tr>
         
          <tr>
           <td style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF;"  >         
			&nbsp;</td>
           <td style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF;"  > <font size="2">        or</td>
           <td style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF;"  >         
			&nbsp;</td>
          </tr>
         
          <tr>
           <td  style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF;"  > <font size="2">
			Create New Group:</td>

           <td  style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF;"  > 

			<asp:TextBox ID="txtNewGroup" runat="Server" Width="294px"></asp:TextBox>&nbsp; <br /> 
</td>
           <td  style="border-style: solid; border-width: 0; text-align: left; background-color: #FFFFFF;"  > 
<asp:Button ID="btnNewGroup" runat="Server" Text="create" /> 
</td>
           
          </tr>
         
         </table> 
		<br />
         <br />
         <br />
         
    <asp:GridView ID="gvSchedule" runat="server" AutoGenerateColumns="False" DataKeyNames="schedule_id"
        DataSourceID="oDSSchedule" CellPadding="5" style="border-collapse: collapse" bordercolor="#111111" BackImageUrl="images/alt_color.gif" Width="600px" AllowSorting="True">
        <Columns>
            <asp:BoundField DataField="schedule_desc" HeaderText="Schedule" SortExpression="schedule_desc" />
            <asp:TemplateField HeaderText="Type" SortExpression="schedule_type">
                <ItemTemplate>
                    <asp:Label ID="lblType" runat="server"></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Days to run" SortExpression="schedule_dow_list">
                
                <ItemTemplate>
                    <asp:Label ID="lblDow" runat="server"></asp:Label>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:BoundField DataField="scheduled_time" SortExpression="scheduled_time" HeaderText="Time" DataFormatString="{0:hh:mm tt}" />
                      
           
            <asp:TemplateField HeaderText="Included">
                <ItemTemplate>
                    <asp:CheckBox ID="chkIncluded" runat="server" />
                    <asp:HiddenField ID="schedule_id" runat="server"/>
                </ItemTemplate>
            </asp:TemplateField>
            
        </Columns>
        <AlternatingRowStyle BackColor="#ffffff" />
    </asp:GridView>
    
    <asp:Button ID="btnUpdate" runat="server" Text="Update" />
    
    <asp:ObjectDataSource ID="oDSSchedule" runat="server" OldValuesParameterFormatString="original_{0}"
        SelectMethod="GetData" TypeName="DataSetProfileSearchScheduleTableAdapters.ScheduleWithGroupIncludeTableAdapter">
        <SelectParameters>
            <asp:SessionParameter Name="user_id" SessionField="user_id" Type="Int32" />
            <asp:ControlParameter ControlID="ddlScheduleGroup" Name="schedule_grp_id" PropertyName="SelectedValue"
                Type="Int32" />
        </SelectParameters>
    </asp:ObjectDataSource>
         
         
         </asp:Content>

