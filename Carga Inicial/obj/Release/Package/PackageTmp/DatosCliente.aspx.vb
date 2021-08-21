Imports ClassLibrary1
Imports System.Web.Hosting
Imports System.Data.SqlClient

Public Class DatosCliente
    Inherits System.Web.UI.Page

    Dim DSEmpresa, DSClientes As DataSet
    Dim SDAEmpresa, SDAClientes As SqlDataAdapter
    Dim DREmpresa, DRClientes As DataRow
    Dim RFC, StrSql, ResultF, Etiq1 As String
    Dim ClaveCli As Integer
    Dim sb As New StringBuilder()
    Dim CliNvo As Boolean
   
    Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim CliNvoTxt = Request.QueryString("CliNvo")
        Dim RFCTxt = Request.QueryString("RFC")
        RFC = RFCTxt
        LblMensajes.Visible = False
        If CargaCliente(RFC) Then
            BtnContinuar.Enabled = True
            CliNvo = False
            CliNvoTxt = "No"
        End If
        If CliNvoTxt = "Si" Then
            CliNvo = True
            BtnContinuar.Enabled = False
            TxtRFC.Text = RFC
            Etiq1 = LblTit1.Text
            LblTit1.Text = "Alta de Cliente"
            LblTit2.Visible = False
            BtnActualiza.Text = "Guardar Datos"
        Else
            CargaTxt()
        End If

    End Sub

    Sub CargaTxt()
        If CargaCliente(RFC) Then
            TxtClaveCli.Text = Trim(DRClientes("Clave"))
            TxtNombre.Text = Trim(DRClientes("nombre"))
            TxtRFC.Text = Trim(DRClientes("rfc"))
            TxtCAlle.Text = Trim(DRClientes("Calle"))
            TxtNumExt.Text = Trim(DRClientes("numexterior"))
            TxtNumInt.Text = Trim(DRClientes("numinterior"))
            TxtColonia.Text = Trim(DRClientes("colonia"))
            TxtCiudad.Text = Trim(DRClientes("localidad"))
            TXtMunicipio.Text = Trim(DRClientes("municipio"))
            TxtEstado.Text = Trim(DRClientes("estado"))
            TxtPais.Text = Trim(DRClientes("Pais"))
            TxtCP.Text = Trim(DRClientes("CP"))
            TxtMail.MaxLength = 250
            TxtMail.Text = Trim(DRClientes("email"))
        Else
            LblMensajes.Text = "La información se actualizó pero no se pudo recargar la página, favor de cerrar sesión y volver a entrar"
        End If
    End Sub

    Protected Sub BtnActualiza_Click(sender As Object, e As EventArgs) Handles BtnActualiza.Click
        If ValidarCliente() Then
            If CliNvo Then
                ClaveCli = GeneraClave()
                If ClaveCli > 0 Then
                    LblMensajes.Text = "Cliente Registrado Exitosamente!!"
                Else
                    LblMensajes.Text = "LO SENTIMOS!! No se pudo generar la clave de Cliente, Inténtelo nuevamente"
                End If
            Else
                LblMensajes.Text = "Datos Modificados Exitosamente!!"
            End If
            StrSql = "Update Clientes set nombre='" + Trim(Request.Form(TxtNombre.ID)).ToUpper + "',Calle='" + Trim(Request.Form(TxtCAlle.ID)) + "',numexterior='" + Request.Form(TxtNumExt.ID) + _
                   "',numinterior='" + Request.Form(TxtNumInt.ID) + "',colonia='" + Trim(Request.Form(TxtColonia.ID)) + "',CP=" + Str(TxtCP.Text) + _
                  ",localidad='" + Trim(Request.Form(TxtCiudad.ID)) + "',municipio='" + Trim(Request.Form(TXtMunicipio.ID)) + "',estado='" + Trim(Request.Form(TxtEstado.ID)) + "',Pais='" + _
                  Trim(Request.Form(TxtPais.ID)) + "',email='" + Trim(Request.Form(TxtMail.ID)) + "' where rfc='" + TxtRFC.Text.ToUpper + "'"
            If ComandoSQL.ComSQL(StrSql) Then
                CargarTablas.BanderaOk = True
                CargaTxt()
            Else
                LblMensajes.Text = "ERROR!! La información no se pudo actualizar, intente nuevamente o llame a su proveedor"
            End If
        End If
        Advertencia()
    End Sub

    Protected Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Response.Redirect("Default.aspx")
    End Sub

    Sub BtnContinuar_Click(sender As Object, e As EventArgs) Handles BtnContinuar.Click
        If GeneraFactura.UserWeb > 20 Then
            GeneraFactura.UserWeb = 0
        Else
            GeneraFactura.UserWeb += 1
        End If
        Dim UserWeb As Integer = GeneraFactura.UserWeb
        ComandoSQL.ComSQL("Delete From TempoFact where Cliente=" + Trim(TxtClaveCli.Text))
        Response.Redirect("DatosTicket.aspx?Clavecliente=" + TxtClaveCli.Text.ToString() + "&RFC=" + TxtRFC.Text.ToString() + "&UserWeb=" + Str(UserWeb))
    End Sub

    Public Function ValidarCliente() As Boolean
        Dim MensajeValida, ValidaT As String
        MensajeValida = ""
        ValidaT = 1
        If Len(TxtNombre.Text) > 0 Then
        Else
            MensajeValida = "Nombre En Blanco." & vbCrLf
        End If

        If Len(Me.TxtRFC.Text) > 0 Then
            If Len(Me.TxtRFC.Text) < 12 Then
                MensajeValida = MensajeValida & "Rfc Longitud Incorrecta" & vbCrLf
            End If
        Else
            MensajeValida = MensajeValida & "Rfc En Blanco." & vbCrLf
        End If

        If Len(Me.TxtCAlle.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Calle En Blanco." & vbCrLf
        End If
        If Len(Me.TxtNumExt.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Numero En Blanco." & vbCrLf
        End If
        If Len(Me.TxtColonia.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Colonia En Blanco." & vbCrLf
        End If
        If Len(Me.TxtCiudad.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Ciudad En Blanco." & vbCrLf
        End If
        If Len(Me.TXtMunicipio.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Municipío En Blanco." & vbCrLf
        End If
        If Len(Me.TxtEstado.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Estado En Blanco." & vbCrLf
        End If
        If Len(Me.TxtPais.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Pais En Blanco." & vbCrLf
        End If
        If Len(Me.TxtCP.Text) > 0 Then
        Else
            MensajeValida = MensajeValida & "Codigo Postal En Blanco." & vbCrLf
        End If
        If Len(MensajeValida) > 0 Then
            ValidarCliente = False
            LblMensajes.Text = "Error en Captura: " + MensajeValida
            CargarTablas.BanderaOk = True
            Advertencia()
        Else
            ValidarCliente = True
        End If
    End Function

    Protected Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOk.Click
        If CargarTablas.BanderaOk Then
            BtnActualiza.Visible = True
            BtnContinuar.Visible = True
            BtnSalir.Visible = True
            BtnOk.Visible = True
            LblMensajes.Visible = False
            BtnOk.Visible = False
            CargarTablas.BanderaOk = False
        Else
            Response.Redirect("Default.aspx")
        End If
    End Sub
    Public Sub Advertencia()
        BtnActualiza.Visible = False
        BtnContinuar.Visible = False
        BtnSalir.Visible = False
        BtnOk.Visible = True
        LblMensajes.Visible = True
    End Sub
    Function CargaCliente(XRFC) As Boolean
        DSClientes = New DataSet
        SDAClientes = New SqlClient.SqlDataAdapter("select * from Clientes where rfc='" + XRFC + "'", CargarTablas.conn)
        SDAClientes.Fill(DSClientes, "Clientes")
        If DSClientes.Tables("Clientes").Rows.Count > 0 Then
            CargaCliente = True
            DRClientes = DSClientes.Tables("Clientes").Rows(0)
        Else
            CargaCliente = False
        End If
    End Function

    Function GeneraClave() As Integer
        DSClientes = New DataSet
        If TxtClaveCli.Text = "" Then
            StrSql = "INSERT INTO Clientes (RFC) values ('" + RFC + "')"
            ComandoSQL.ComSQL(StrSql)
        End If
        SDAClientes = New SqlClient.SqlDataAdapter("select * from Clientes where RFC='" + RFC + "'", CargarTablas.conn)
        SDAClientes.Fill(DSClientes, "Clientes")
        If DSClientes.Tables("Clientes").Rows.Count > 0 Then
            DRClientes = DSClientes.Tables("Clientes").Rows(0)
            StrSql = "update Clientes set clave=id where id=" + Str(DRClientes("Id"))
            ComandoSQL.ComSQL(StrSql)
            GeneraClave = DRClientes("Id")
        Else
            GeneraClave = 0
        End If
    End Function
End Class