<%@ Page Language="VB" MasterPageFile="~/Admin.master" AutoEventWireup="false" CodeFile="RezCentralHeader.aspx.vb" Inherits="RezCentralHeader" title="Untitled Page" %>
<%@ Register TagPrefix="rh" TagName="DollarPercent" Src="~/DollarPercent.ascx" %>
<asp:Content ID="Content1" ContentPlaceHolderID="cphMain" Runat="Server">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
    <td align="left"><asp:Label ID="lblMessage" runat="server"></asp:Label></td>
    <td align="right"><asp:HyperLink ID="lnkBack" runat="Server" Text="Back to Previous Page"></asp:HyperLink></td>
</tr>
</table>

    <br />
       <asp:ValidationSummary ID="vs1" runat="server" />
    <asp:GridView ID="GridView1" runat="server" DataSourceID="odsRezHeader" AutoGenerateColumns="False" DataKeyNames="RezCentralHeaderID" HeaderStyle-Font-Bold="false" ShowFooter="True">
        <Columns>
            <asp:BoundField DataField="RezCentralHeaderID" HeaderText="RezCentralHeaderID" InsertVisible="False"
                ReadOnly="True" SortExpression="RezCentralHeaderID" Visible="False" />
            <asp:TemplateField HeaderText="Branch" SortExpression="Branch">
                <EditItemTemplate>
                    <asp:TextBox ID="txtBranchEdit" Width="50px" runat="server" Text='<%# Bind("Branch") %>'></asp:TextBox>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label1" Width="50px" runat="server" Text='<%# Bind("Branch") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtBranch" Width="50px"  runat="server" Text='<%# Bind("Branch") %>'></asp:TextBox>
                </FooterTemplate>
                
            </asp:TemplateField>
            
            <asp:TemplateField HeaderText="Rate Code" SortExpression="RateCode">
                <EditItemTemplate>
                    <asp:TextBox ID="txtRateCodeEdit" Width="80px" runat="server" Text='<%# Bind("RateCode") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4Edit" Display="None" ControlToValidate="txtRateCodeEdit" runat="server" Text="*" ErrorMessage="Rate Code is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label4" Width="80px" runat="server" Text='<%# Bind("RateCode") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtRateCode" Width="80px" runat="server" Text='<%# Bind("RateCode") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator4" Display="None" ControlToValidate="txtRateCode" runat="server" Text="*" ErrorMessage="Rate Code is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="TsdSystem" SortExpression="TsdSystem" >
                <EditItemTemplate>
                    <asp:DropDownList CssClass="nugrid_input" ID="ddlSystem" runat="Server" DataSourceID="oDSSystems" DataTextField="tsd_system" DataValueField="tsd_system" SelectedValue='<%# Bind("TsdSystem") %>' ></asp:DropDownList>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:Label ID="lblTsdSystem" Text='<%# Bind("TsdSystem") %>' runat="Server"></asp:Label>                    
                </ItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList CssClass="nugrid_input" ID="ddlSystem" runat="Server" DataSourceID="oDSSystems" DataTextField="tsd_system" DataValueField="tsd_system" ></asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Dly Diff" SortExpression="DailyDiff">
                <EditItemTemplate>
                    <asp:TextBox ID="txtDailyDiffEdit" Width="35px" runat="server" Text='<%# Bind("DailyDiff", "{0:F2}") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5Edit" Display="none" ControlToValidate="txtDailyDiffEdit" runat="server" Text="*" ErrorMessage="Daily Diff is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label5" runat="server" Text='<%# Bind("DailyDiff", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtDailyDiff" Width="35px" runat="server" Text='0.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator5" Display="none" ControlToValidate="txtDailyDiff" runat="server" Text="*" ErrorMessage="Daily Diff is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator5" Display="none" ControlToValidate="txtDailyDiff" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Daily Diff must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
                
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Is $" SortExpression="DailyIsDollar">
                <EditItemTemplate>
                    <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# Bind("DailyIsDollar") %>' />
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# Bind("DailyIsDollar") %>'
                        Enabled="false" />
                </ItemTemplate>
                <FooterTemplate>
                <asp:CheckBox ID="chkDailyIsDollar" runat="server" Checked='true'
                        />
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Wknd Diff" SortExpression="WeekendDiff">
                <EditItemTemplate>
                    <asp:TextBox ID="txtWeekendDiffEdit" Width="35px"  runat="server" Text='<%# Bind("WeekendDiff", "{0:F2}") %>'></asp:TextBox>
<asp:RequiredFieldValidator ID="RequiredFieldValidator6Edit" ControlToValidate="txtWeekendDiffEdit" Display="None" runat="server" Text="*" ErrorMessage="Weekend Diff is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ControlStyle CssClass="nugrid_input" Width="35px" />
                <ItemTemplate>
                    <asp:Label ID="Label6" Width="35px" runat="server" Text='<%# Bind("WeekendDiff", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtWeekendDiff" Width="35px" runat="server" Text='0.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator6" Display="none" ControlToValidate="txtWeekendDiff" runat="server" Text="*" ErrorMessage="Weekend Diff is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator6" Display="none" ControlToValidate="txtWeekendDiff" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Weekend Diff must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
                
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Is $" SortExpression="WeekendIsDollar">
                <EditItemTemplate>
                    <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# Bind("WeekendIsDollar") %>' />
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:CheckBox ID="chkWeekendIsDollar" runat="server" Checked='<%# Bind("WeekendIsDollar") %>'
                        Enabled="false" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:CheckBox ID="chkWeekendIsDollar" runat="server" Checked='true' />
                </FooterTemplate>
                
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Wkly Diff" SortExpression="WeeklyDiff">
                <EditItemTemplate>
                    <asp:TextBox ID="txtWeeklyDiffEdit" Width="35px"  runat="server" Text='<%# Bind("WeeklyDiff", "{0:F2}") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator7Edit" ControlToValidate="txtWeeklyDiffEdit" Display="None" runat="server" Text="*" ErrorMessage="Weekly Diff is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ControlStyle CssClass="nugrid_input" Width="35px" />
                <ItemTemplate>
                    <asp:Label ID="Label7" runat="server" Text='<%# Bind("WeeklyDiff", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtWeeklyDiff" Width="35px"  runat="server" Text='0.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator7" Display="none" ControlToValidate="txtWeeklyDiff" runat="server" Text="*" ErrorMessage="Weekly Diff is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator7" Display="none" ControlToValidate="txtWeeklyDiff" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Weekly Diff must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
                
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Is $" SortExpression="WeeklyIsDollar">
                <EditItemTemplate>
                    <asp:CheckBox ID="CheckBox3" runat="server" Checked='<%# Bind("WeeklyIsDollar") %>' />
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:CheckBox ID="CheckBox3" runat="server" Checked='<%# Bind("WeeklyIsDollar") %>'
                        Enabled="false" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:CheckBox ID="chkWeeklyIsDollar" runat="server" Checked='true' />
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Mthly Diff" SortExpression="MonthlyDiff">
                <EditItemTemplate>
                    <asp:TextBox ID="txtMonthlyDiffEdit" Width="35px" runat="server" Text='<%# Bind("MonthlyDiff", "{0:F2}") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8Edit" ControlToValidate="txtMonthlyDiffEdit" Display="None" runat="server" Text="*" ErrorMessage="Monthly Diff is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ControlStyle CssClass="nugrid_input" Width="35px" />
                <ItemTemplate>
                    <asp:Label ID="Label8" runat="server" Text='<%# Bind("MonthlyDiff", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtMonthlyDiff" Width="35px" runat="server" Text='0.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator8" Display="none" ControlToValidate="txtMonthlyDiff" runat="server" Text="*" ErrorMessage="Monthly Diff is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator8" Display="none" ControlToValidate="txtMonthlyDiff" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Monthly Diff must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Is $" SortExpression="MonthlyIsDollar">
                <EditItemTemplate>
                    <asp:CheckBox ID="CheckBox4" runat="server" Checked='<%# Bind("MonthlyIsDollar") %>' />
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:CheckBox ID="CheckBox4" runat="server" Checked='<%# Bind("MonthlyIsDollar") %>'
                        Enabled="false" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:CheckBox ID="chkMonthlyIsDollar" runat="server" Checked='true' />
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="XDay Diff" SortExpression="ExtraDayDiff">
                <EditItemTemplate>
                    <asp:TextBox ID="txtExtraDayDiffEdit" Width="35px" runat="server" Text='<%# Bind("ExtraDayDiff", "{0:F2}") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator9Edit" Display="none" ControlToValidate="txtExtraDayDiffEdit" runat="server" Text="*" ErrorMessage="Extra Day Diff is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ControlStyle Width="35px" />
                <ItemTemplate>
                    <asp:Label ID="Label9" runat="server" Text='<%# Bind("ExtraDayDiff", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtExtraDayDiff" Width="35px" runat="server" Text='0.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator9" Display="none" ControlToValidate="txtExtraDayDiff" runat="server" Text="*" ErrorMessage="Extra Day Diff is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator9" Display="none" ControlToValidate="txtExtraDayDiff" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Extra Day Diff must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Is $" SortExpression="ExtraDayIsDollar">
                <EditItemTemplate>
                    <asp:CheckBox ID="CheckBox5" runat="server" Checked='<%# Bind("ExtraDayIsDollar") %>' />
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:CheckBox ID="CheckBox5" runat="server" Checked='<%# Bind("ExtraDayIsDollar") %>'
                        Enabled="false" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:CheckBox ID="chkExtraDayIsDollar" runat="server" Checked='true' />
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="WkndX Diff" SortExpression="WkndExtraDayDiff">
                <EditItemTemplate>
                    <asp:TextBox ID="txtWkndExtraDayDiffEdit" Width="35px" runat="server" Text='<%# Bind("WkndExtraDayDiff", "{0:F2}") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator10Edit" Display="none" ControlToValidate="txtWkndExtraDayDiffEdit" runat="server" Text="*" ErrorMessage="Weekend Extra Day Diff is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ControlStyle Width="30px" />
                <ItemTemplate>
                    <asp:Label ID="Label10" runat="server" Text='<%# Bind("WkndExtraDayDiff", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtWkndExtraDayDiff" Width="35px" runat="server" Text='0.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator10" Display="none" ControlToValidate="txtWkndExtraDayDiff" runat="server" Text="*" ErrorMessage="Weekend Extra Day Diff is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator10" Display="none" ControlToValidate="txtWkndExtraDayDiff" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Weekend Extra Day Diff must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Is $" SortExpression="WkndExtraDayIsDollar">
                <EditItemTemplate>
                    <asp:CheckBox ID="CheckBox6" runat="server" Checked='<%# Bind("WkndExtraDayIsDollar") %>' />
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:CheckBox ID="CheckBox6" runat="server" Checked='<%# Bind("WkndExtraDayIsDollar") %>'
                        Enabled="false" />
                </ItemTemplate>
                <FooterTemplate>
                    <asp:CheckBox ID="chkWkndExtraDayIsDollar" runat="server" Checked='true' />
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="WklyX Factor" SortExpression="WeeklyExtraDayFactor">
                <EditItemTemplate>
                    <asp:TextBox ID="txtWeeklyExtraDayFactorEdit" Width="25px" runat="server" Text='<%# Bind("WeeklyExtraDayFactor", "{0:F2}") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator11Edit" Display="None"  ControlToValidate="txtWeeklyExtraDayFactorEdit" runat="server" Text="*" ErrorMessage="Weekly Extra Day Factor is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ControlStyle CssClass="nugrid_input" Width="25px" />
                <ItemTemplate>
                    <asp:Label ID="Label11" runat="server" Text='<%# Bind("WeeklyExtraDayFactor", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtWeeklyExtraDayFactor" Width="25px" runat="server" Text='5.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator11" Display="none"  ControlToValidate="txtWeeklyExtraDayFactor" runat="server" Text="*" ErrorMessage="Weekly Extra Day Factor is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator11" Display="none" ControlToValidate="txtWeeklyExtraDayFactor" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Weekly Extra Day Factor must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
                
            </asp:TemplateField>
            <asp:TemplateField HeaderText="MthlyX Factor" SortExpression="MonthlyExtraDayFactor">
                <EditItemTemplate>
                    <asp:TextBox ID="txtMonthlyExtraDayFactorEdit" Width="30px" runat="server" Text='<%# Bind("MonthlyExtraDayFactor", "{0:F2}") %>'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator12Edit" Display="None" ControlToValidate="txtMonthlyExtraDayFactorEdit" runat="server" Text="*" ErrorMessage="Monthly Extra Day Factor is a required field." ValidationGroup="Edit"></asp:RequiredFieldValidator>
                </EditItemTemplate>
                <ControlStyle CssClass="nugrid_input" Width="30px" />
                <ItemTemplate>
                    <asp:Label ID="Label12" runat="server" Text='<%# Bind("MonthlyExtraDayFactor", "{0:F2}") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtMonthlyExtraDayFactor" Width="30px" runat="server" Text='21.00'></asp:TextBox>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator12" ControlToValidate="txtMonthlyExtraDayFactor" runat="server" Text="*" ErrorMessage="Monthly Extra Day Factor is a required field." ValidationGroup="Insert"></asp:RequiredFieldValidator>
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator12" ControlToValidate="txtMonthlyExtraDayFactor" runat="server" ValidationExpression="^\d*\.?\d*$" Text="*" ErrorMessage="Monthly Extra Day Factor must be a decimal."></asp:RegularExpressionValidator>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="DOW" SortExpression="WeekendDOW">
                <EditItemTemplate>
                    <asp:TextBox ID="TextBox13" Width="50px" runat="server" Text='<%# Bind("WeekendDOW") %>'></asp:TextBox>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:Label ID="Label13" runat="server" Text='<%# Bind("WeekendDOW") %>'></asp:Label>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:TextBox ID="txtWeekendDOW" Width="50px" runat="server" Text='<%# Bind("WeekendDOW") %>'></asp:TextBox>
                </FooterTemplate>
            </asp:TemplateField>
            
            <asp:TemplateField HeaderText="Type" SortExpression="GovType">
                <EditItemTemplate>
                    <asp:DropDownList CssClass="nugrid_input" ID="ddlGovType" runat="Server" SelectedValue='<%# Bind("GovType") %>' >
                        <asp:ListItem Value="0" Text="Regular"></asp:ListItem>
                        <asp:ListItem Value="1" Text="Gov Rounded"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Gov Truncated"></asp:ListItem>
                    </asp:DropDownList>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:Label ID="lblGovType" runat="Server" ></asp:Label>
                    <asp:DropDownList visible="false" CssClass="nugrid_input" ID="ddlGovType" runat="Server" SelectedValue='<%# Bind("GovType") %>' >
                        <asp:ListItem Value="0" Text="Regular"></asp:ListItem>
                        <asp:ListItem Value="1" Text="Gov Rounded"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Gov Truncated"></asp:ListItem>
                    </asp:DropDownList>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:DropDownList CssClass="nugrid_input" ID="ddlGovType" runat="Server" SelectedValue='<%# Bind("GovType") %>' >
                        <asp:ListItem Value="0" Text="Regular"></asp:ListItem>
                        <asp:ListItem Value="1" Text="Gov Rounded"></asp:ListItem>
                        <asp:ListItem Value="2" Text="Gov Truncated"></asp:ListItem>
                    </asp:DropDownList>
                </FooterTemplate>
            </asp:TemplateField>
            <asp:TemplateField ShowHeader="False">
                <EditItemTemplate>
                    <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Update" ValidationGroup="Edit"
                        Text="Update"></asp:LinkButton>
                    <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel"
                        Text="Cancel"></asp:LinkButton>
                </EditItemTemplate>
                <ItemTemplate>
                    <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandName="Edit" ValidationGroup="Edit"
                        Text="Edit"></asp:LinkButton>
                    <asp:LinkButton ID="btnDelete" runat="server" CausesValidation="False" CommandName="Delete"
                        Text="Delete"></asp:LinkButton>
                </ItemTemplate>
                <FooterTemplate>
                    <asp:LinkButton ID="btnInsert" runat="server" CausesValidation="True" CommandName="Insert" ValidationGroup="Insert"
                        Text="Insert"></asp:LinkButton>
                    
                </FooterTemplate>
            </asp:TemplateField>
            
        </Columns>
        
        <HeaderStyle CssClass="nugrid_header" Font-Bold="False" />
        <RowStyle CssClass="nugrid_light" />
        <AlternatingRowStyle CssClass="nugrid_dark" />
        
    </asp:GridView>
    <br /><br />
    <table width="150" border="1" cellpadding="3" cellspacing="0">
        <tr><td colspan=2>DOW Legend</td></tr>
        <tr><td>1</td><td>Sunday</td></tr>
        <tr><td>2</td><td>Monday</td></tr>
        <tr><td>3</td><td>Tuesday</td></tr>
        <tr><td>4</td><td>Wednesday</td></tr>
        <tr><td>5</td><td>Thurday</td></tr>
        <tr><td>6</td><td>Friday</td></tr>
        <tr><td>7</td><td>Saturday</td></tr>
        
    </table>
    <asp:ObjectDataSource ID="oDSSystems" runat="server" OldValuesParameterFormatString="original_{0}"
                    SelectMethod="GetData" TypeName="DataSetLookupsTableAdapters.rezcentral_systemsTableAdapter">
                </asp:ObjectDataSource>
    <asp:ObjectDataSource ID="odsRezHeader" runat="server" DeleteMethod="Delete" InsertMethod="Insert"
        OldValuesParameterFormatString="original_{0}" SelectMethod="GetDataByOrg" TypeName="DataSetRezTableAdapters.newrez_headerTableAdapter"
        UpdateMethod="Update">
        <DeleteParameters>
            <asp:Parameter Name="original_RezCentralHeaderID" Type="Int32" />
        </DeleteParameters>
        <UpdateParameters>
            <asp:Parameter Name="Branch" Type="String" />
            <asp:Parameter Name="RateCode" Type="String" />
            <asp:Parameter Name="TsdSystem" Type="String" />
            <asp:Parameter Name="DailyDiff" Type="Decimal" />
            <asp:Parameter Name="DailyIsDollar" Type="Boolean" />
            <asp:Parameter Name="WeekendDiff" Type="Decimal" />
            <asp:Parameter Name="WeekendIsDollar" Type="Boolean" />
            <asp:Parameter Name="WeeklyDiff" Type="Decimal" />
            <asp:Parameter Name="WeeklyIsDollar" Type="Boolean" />
            <asp:Parameter Name="MonthlyDiff" Type="Decimal" />
            <asp:Parameter Name="MonthlyIsDollar" Type="Boolean" />
            <asp:Parameter Name="ExtraDayDiff" Type="Decimal" />
            <asp:Parameter Name="ExtraDayIsDollar" Type="Boolean" />
            <asp:Parameter Name="WkndExtraDayDiff" Type="Decimal" />
            <asp:Parameter Name="WkndExtraDayIsDollar" Type="Boolean" />
            <asp:Parameter Name="WeeklyExtraDayFactor" Type="Decimal" />
            <asp:Parameter Name="MonthlyExtraDayFactor" Type="Decimal" />
            <asp:Parameter Name="WeekendDOW" Type="String" DefaultValue="" />
            <asp:Parameter Name="GovType" Type="Byte" />
            <asp:Parameter Name="original_RezCentralHeaderID" Type="Int32" />
        </UpdateParameters>
        <SelectParameters>
            <asp:SessionParameter Name="org_id" SessionField="org_id" Type="Int32" />
        </SelectParameters>
        <InsertParameters>
            <asp:SessionParameter Name="org_id" Type="Int32" SessionField="org_id" />
            <asp:Parameter Name="Branch" Type="String" />
            <asp:Parameter Name="SenderID" Type="String" />
            <asp:Parameter Name="RecipientID" Type="String" />
            <asp:Parameter Name="TradingPartnerCode" Type="String" />
            <asp:Parameter Name="MessageID" Type="String" />
            <asp:Parameter Name="RateCode" Type="String" />
            <asp:Parameter Name="TsdSystem" Type="String" />
            <asp:Parameter Name="DailyDiff" Type="Decimal" />
            <asp:Parameter Name="DailyIsDollar" Type="Boolean" />
            <asp:Parameter Name="WeekendDiff" Type="Decimal" />
            <asp:Parameter Name="WeekendIsDollar" Type="Boolean" />
            <asp:Parameter Name="WeeklyDiff" Type="Decimal" />
            <asp:Parameter Name="WeeklyIsDollar" Type="Boolean" />
            <asp:Parameter Name="MonthlyDiff" Type="Decimal" />
            <asp:Parameter Name="MonthlyIsDollar" Type="Boolean" />
            <asp:Parameter Name="ExtraDayDiff" Type="Decimal" />
            <asp:Parameter Name="ExtraDayIsDollar" Type="Boolean" />
            <asp:Parameter Name="WkndExtraDayDiff" Type="Decimal" />
            <asp:Parameter Name="WkndExtraDayIsDollar" Type="Boolean" />
            <asp:Parameter Name="WeeklyExtraDayFactor" Type="Decimal" />
            <asp:Parameter Name="MonthlyExtraDayFactor" Type="Decimal" />
            <asp:Parameter Name="WeekendDOW" Type="String" />
            <asp:Parameter Name="GovType" Type="Byte" />
        </InsertParameters>
    </asp:ObjectDataSource>
    &nbsp;
    
</asp:Content>

