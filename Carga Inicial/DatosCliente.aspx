<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="DatosCliente.aspx.vb" Inherits="FactuGasWEB.DatosCliente" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
        <p style="width: 1017px; height: 14px; margin-left: 40px">
&nbsp;
            <asp:Image ID="Image1" runat="server" Height="120px" ImageUrl="Logo.jpg" Width="120px" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Image ID="Image2" runat="server" Height="120px" ImageUrl="Totalv.jpg" Width="120px" />
        </p>
        <p style="width: 1020px; margin-left: 40px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <br />
            <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="LblTit1" runat="server" Font-Bold="True" Font-Size="X-Large" ForeColor="Blue" Text="FAVOR DE VERIFICAR SU INFORMACIÓN"></asp:Label>
            <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;
            <asp:Label ID="LblTit2" runat="server" Font-Bold="True" Font-Size="X-Large" ForeColor="Blue" Text="O MODIFIQUE Y PULSE ACTUALIZAR DATOS"></asp:Label>
            <br />
            <br />
            <asp:Label ID="Label14" runat="server" Font-Bold="True" Text="Clave Cliente: "></asp:Label>
&nbsp;<asp:TextBox ID="TxtClaveCli" runat="server" style="text-transform:uppercase" Width="43px" BackColor="#CCCCCC" Font-Bold="True" ReadOnly="True"></asp:TextBox>
            <br />
            <br />
            <asp:Label ID="Label1" runat="server" Font-Bold="True" Text="RFC"></asp:Label>
&nbsp;:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtRFC" runat="server" style="text-transform:uppercase" Width="175px" BackColor="#CCCCCC" Font-Bold="True" ReadOnly="True"></asp:TextBox>
&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label2" runat="server" Font-Bold="True" Text="Nombre:"></asp:Label>
&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtNombre" runat="server" style="text-transform:uppercase" Width="621px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
            <br />
            <br />
            <asp:Label ID="Label6" runat="server" Font-Bold="True" Text="Calle:"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtCAlle" runat="server" style="text-transform:uppercase" Width="538px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label4" runat="server" Font-Bold="True" Text="Num. Ext."></asp:Label>
&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtNumExt" runat="server" Width="50px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label5" runat="server" Font-Bold="True" Text="Num. Int."></asp:Label>
&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtNumInt" runat="server" Width="50px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
            <br />
            <br />
            <asp:Label ID="Label7" runat="server" Font-Bold="True" Text="Colonia o Fracc:"></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtColonia" runat="server" style="text-transform:uppercase" Width="538px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
            <br />
            <br />
            <asp:Label ID="Label8" runat="server" Font-Bold="True" Text="Localidad o Municipio:"></asp:Label>
&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TXtMunicipio" runat="server" style="text-transform:uppercase" Width="538px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
            <br />
            <br />
            <asp:Label ID="Label9" runat="server" Font-Bold="True" Text="Ciudad:"></asp:Label>
&nbsp;
            <asp:TextBox ID="TxtCiudad" runat="server" style="text-transform:uppercase" Width="411px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label10" runat="server" Font-Bold="True" Text="Estado: "></asp:Label>
&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtEstado" runat="server" style="text-transform:uppercase" Width="411px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
            <br />
            <br />
            <asp:Label ID="Label11" runat="server" Font-Bold="True" Text="Código Postal: "></asp:Label>
&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtCP" runat="server" Width="58px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label12" runat="server" Font-Bold="True" Text="País:"></asp:Label>
&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtPais" runat="server" style="text-transform:uppercase"  Width="175px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
&nbsp;&nbsp;&nbsp;
            <asp:Label ID="Label13" runat="server" Font-Bold="True" Text="Email para recibir Facturas:"></asp:Label>
&nbsp;&nbsp;&nbsp;
            <asp:TextBox ID="TxtMail" runat="server" Width="332px" Font-Bold="True" ForeColor="#0000CC"></asp:TextBox>
            <br />
            <br />
            <br />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Button ID="BtnActualiza" runat="server" Font-Bold="True" Font-Size="Medium" Height="43px" Text="Actualizar Datos" Width="156px" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Button ID="BtnContinuar" runat="server" Font-Bold="True" Font-Size="Medium" Height="43px" Text="Continuar" Width="156px" />
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:Button ID="BtnSalir" runat="server" Font-Bold="True" Font-Size="Medium" Height="43px" Text="Cerrar Sesión" Width="156px" />
            <br />
            <strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <asp:Label ID="LblMensajes" runat="server" Font-Size="X-Large" ForeColor="#CC3300">RFC No encontrado!! desea registrarse?</asp:Label>
            &nbsp;&nbsp;
            <asp:Button ID="BtnOk" runat="server" Text="Ok" Visible="False" />
            </strong>
            <br />
            <br />
            <br />
        </p>
    
    </div>
    </form>
</body>
</html>
