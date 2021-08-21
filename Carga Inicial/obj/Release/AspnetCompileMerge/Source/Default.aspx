<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="WebApplication1._Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <style type="text/css">

        .auto-style1 {
            color: #0000FF;
        }
        </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <h1 class="auto-style1" style="width: 787px; margin-top: 0px; height: 146px;">&nbsp;&nbsp;
            <asp:Image ID="Image1" runat="server" Height="150px" ImageUrl="Logo.jpg" Width="150px" />
&nbsp;&nbsp;
            <asp:Label ID="Label3" runat="server" Font-Size="XX-Large" Text="Facturación Web AdminGas"></asp:Label>
&nbsp;
            <asp:Image ID="Image2" runat="server" Height="150px" ImageUrl="Totalv.jpg" Width="150px" />
        </h1>
        <h1 class="auto-style1" style="width: 785px; margin-top: 0px; height: 38px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label2" runat="server" Font-Italic="True" Font-Size="X-Large" ForeColor="Black" Text="Introduzca su RFC"></asp:Label>
            <strong><em>&nbsp; </em>
            <asp:TextBox ID="TxtRFC" runat="server" style="text-transform:uppercase"  Font-Bold="True" Font-Size="Medium" Height="20px" Width="187px"></asp:TextBox>
&nbsp;&nbsp; </strong></h1>
        <h1 class="auto-style1" style="width: 784px; margin-top: 0px; height: 26px;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>&nbsp;</strong><asp:Button ID="BtnIngreso" runat="server" Font-Bold="True" Height="39px" Text="Ingresar" Width="129px" />
            <br />
            <br />
            <br />
        </h1>
        <p class="auto-style1" style="width: 784px; margin-top: 0px; height: 26px;">&nbsp;</p>
        <h1 class="auto-style1" style="width: 716px; margin-top: 0px; height: 20px; margin-left: 80px;">&nbsp; <asp:Label ID="LblEstacion" runat="server" Font-Bold="True" Font-Size="Medium" ForeColor="#CC0000" Text="Label"></asp:Label>
            <br />
        </h1>
    
    </div>
    </form>
</body>
</html>
