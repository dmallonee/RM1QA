<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default3.aspx.vb" Inherits="Default3" %>

<%@ Register Assembly="ComponentArt.Web.UI" Namespace="ComponentArt.Web.UI" TagPrefix="ComponentArt" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>

<script language="javascript" type="text/javascript">
function CheckItems()
{ 
var grid=Grid1; // cagrid is the id of the componentart grid

var gridItem;
var itemIndex = 0;
var strIDs = "";
var checked = false;

while(gridItem = grid.Table.GetRow(itemIndex))
{
if (gridItem.Cells[0].Value) // 0 is the 1st column as it is of checkbox type in this case
{ 
strIDs = strIDs + "," +  gridItem.Cells[1].Value;
}

itemIndex++;
} 

/*if (!checked)
{
alert(’You have not selected any items.\n Please select atleast one item.’);
}*/

document.form1.txt.value = strIDs.substr(1,strIDs.length);
//alert(strIDs);
//return strIDs;
}


</script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table id="Table1" cellpadding="2" cellspacing="2" style="z-index: 101; left: 13px;
            position: absolute; top: 57px">
            <tr>
                <td align="right">
                    <strong>Connection String: </strong>
                </td>
                <td>
                    <asp:TextBox ID="txtConn" runat="server" CssClass="StdTextBox" Text=" DATABASE=Northwind;SERVER= localhost;UID=sa;PWD=;"
                        Width="471px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="right">
                    <strong>SQL string: </strong>&lt;
                </td>
                <td>
                    <asp:TextBox ID="txtSQL" runat="server" CssClass="StdTextBox" Text="DATABASE=Northwind;SERVER=localhost;UID=sa;PWD=;"
                        Width="471px">Select EmployeeId, FirstName, LastName, City from Employees</asp:TextBox></td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    <asp:Button ID="btnLoad" runat="server" BackColor="Silver" BorderStyle="Outset" BorderWidth="2px"
                        Font-Names="Verdana" Font-Size="8pt" Height="22px" Text="Load" Width="59px" />
                    <componentart:grid id="Grid1" runat="server" width="500">
      <Levels>
                        <componentart:GridLevel>
                                           <Columns>             
               <ComponentArt:GridColumn AllowEditing="True" Visible="true" ColumnType="CheckBox" />
               <ComponentArt:GridColumn DataField="EmployeeID"  IsSearchable="True"  HeadingCellCssClass="FirstHeadingCell" DataCellCssClass="FirstDataCell"  HeadingText="EmployeeID"  AllowGrouping="False"  Width="300" />
               <ComponentArt:GridColumn DataField="FirstName"     IsSearchable="True"  HeadingCellCssClass="HeadingCell" DataCellCssClass="DataCell"  HeadingText="First Name" Width="400" />
               <ComponentArt:GridColumn DataField="LastName"     IsSearchable="True"  Width="300"  HeadingText="Last Name"/>
               <ComponentArt:GridColumn DataField="City"     IsSearchable="True"  Width="200" HeadingText="City" />
             </Columns>
                        
                        </componentart:GridLevel>
                        </Levels>   
                    
                    </componentart:grid>
                </td>
            </tr>
            <tr>
                <td colspan="2">
                    <asp:Button ID="btnShow" runat="server" BackColor="Silver" BorderStyle="Outset" BorderWidth="2px"
                        Font-Names="Verdana" Font-Size="8pt" Text="Show me my selection" Visible="False" />&nbsp;
                    <asp:TextBox ID="txt" runat="server"></asp:TextBox></td>
            </tr>
            <tr>
                <td colspan="2">
                    </td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
