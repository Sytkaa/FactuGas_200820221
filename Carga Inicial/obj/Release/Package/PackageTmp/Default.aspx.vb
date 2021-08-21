Imports System.Web.Hosting
Imports ClassLibrary1
Imports System.Data.SqlClient
Imports System.Web

Public Class _Default
    Inherits System.Web.UI.Page

    Dim DSClientes As DataSet
    Dim SDAClientes As SqlDataAdapter
    Dim DRClientes As DataRow
    Dim QueryCarga, RespBox As String
    Dim conn As New SqlClient.SqlConnection

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        BtnSi.Visible = False
        BtnNo.Visible = False
        LblRegistro.Visible = False
        CargarTablas.Path = HostingEnvironment.ApplicationPhysicalPath
        If CargarTablas.Conectar() Then
            CargarTablas.CargaEmpresa()
            LblPAC.Text = CargarTablas.DREmpresa("XTIMBRADO")
            LblEstacion.Text = CargarTablas.DREmpresa("NombreCorto")
        Else
            LblRegistro.Text = "ERROR!! No se pudo conectar a la base de datos, favor de llamar a la Estación!!"
        End If
    End Sub

    Protected Sub BtnIngreso_Click(sender As Object, e As EventArgs) Handles BtnIngreso.Click
        TxtRFC.Text = TxtRFC.Text.ToUpper
        If Len(Trim(TxtRFC.Text)) < 10 Then
            TxtRFC.Visible = False
            BtnIngreso.Visible = False
            LblRegistro.Text = "RFC INVÁLIDO, Favor de ingresar un RFC Válido"
            BtnSi.Text = "Ok"
            LblRegistro.Visible = True
            BtnSi.Visible = True
        Else

            If CargaCliente() Then
                ComandoSQL.ComSQL("Delete From TempoFact where Cliente=" + Trim(Str(DRClientes("Clave"))))
                Response.Redirect("DatosCliente.aspx?CliNvo=" + BtnNo.Text.ToString() + "&RFC=" + TxtRFC.Text.ToString())
            Else
                BtnSi.Visible = True
                BtnNo.Visible = True
                LblRegistro.Visible = True
                BtnIngreso.Visible = False
                TxtRFC.Visible = False
            End If
        End If
    End Sub


    Protected Sub BtnSi_Click(sender As Object, e As EventArgs) Handles BtnSi.Click
        If BtnSi.Text = "Si" Then
            Response.Redirect("DatosCliente.aspx?CliNvo=" + BtnSi.Text.ToString() + "&RFC=" + TxtRFC.Text.ToString())
        Else
            Response.Redirect("Default.aspx")
        End If
    End Sub

    Protected Sub BtnNo_Click(sender As Object, e As EventArgs) Handles BtnNo.Click
        Response.Redirect("Default.aspx")
    End Sub

    Function CargaCliente() As Boolean
        DSClientes = New DataSet
        SDAClientes = New SqlClient.SqlDataAdapter("select * from Clientes where rfc='" + Trim(TxtRFC.Text) + "'", CargarTablas.conn)
        SDAClientes.Fill(DSClientes, "Clientes")
        If DSClientes.Tables("Clientes").Rows.Count > 0 Then
            CargaCliente = True
            DRClientes = DSClientes.Tables("Clientes").Rows(0)
        Else
            CargaCliente = False
        End If
    End Function

    Protected Sub TxtRFC_TextChanged(sender As Object, e As EventArgs) Handles TxtRFC.TextChanged
        TxtRFC.Text.ToUpper()
    End Sub

    Protected Sub BtnDescargar_Click(sender As Object, e As EventArgs) Handles BtnDescargar.Click
        TxtRFC.Text = TxtRFC.Text.ToUpper
        If Len(Trim(TxtRFC.Text)) < 10 Then
            TxtRFC.Visible = False
            BtnIngreso.Visible = False
            BtnDescargar.Visible = False
            LblRegistro.Text = "RFC INVÁLIDO, Favor de ingresar un RFC Válido"
            BtnSi.Text = "Ok"
            LblRegistro.Visible = True
            BtnSi.Visible = True
        Else
            If CargaCliente() Then
                Response.Redirect("DescarFacts.aspx?RFC=" + TxtRFC.Text)
            Else
                LblRegistro.Text = "RFC NO ENCONTRADO, Favor de ingresar de nuevo"
                BtnSi.Text = "Ok"
            End If
        End If

    End Sub
End Class