Imports ClassLibrary1
Imports System.Data.SqlClient

Public Class DescarFacts
    Inherits System.Web.UI.Page
    Dim RFC, StrSql, ValText As String
    Dim Espacios As Integer
    Dim DSEmpresa, DSClientes, DSFacturas As DataSet
    Dim SDAEmpresa, SDAClientes, SDAFacturas As SqlDataAdapter
    Dim DREmpresa, DRClientes, DRFacturas As DataRow
    Dim ClaveCli As Integer
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        RFC = Request.QueryString("RFC")
        Cargacliente()
        LblNombre.Text = Trim(DRClientes("Nombre"))
        If CargaFacturas() Then
            Dim Facturas As New ArrayList
            For i = 0 To DSFacturas.Tables("Facturas").Rows.Count - 1
                DRFacturas = DSFacturas.Tables("Facturas").Rows(i)
                Espacios = (10 - Len(Str(DRFacturas("Folioint"))))
                ValText = DRFacturas("Fecha") + New String("_"c, Espacios) + Str(DRFacturas("Folioint"))
                Espacios = 15 - Len(FormatCurrency(DRFacturas("Total")))
                ValText = ValText + New String("_"c, Espacios) + FormatCurrency(DRFacturas("Total"))
                Facturas.Add(ValText)
            Next
            LBxFacts.DataSource = Facturas
            LBxFacts.DataBind()
        Else
            Response.Redirect("NoFacturas.aspx")
        End If

    End Sub

    Sub Cargacliente()
        DSClientes = New DataSet
        SDAClientes = New SqlClient.SqlDataAdapter("select * from Clientes where rfc='" + RFC + "'", CargarTablas.conn)
        SDAClientes.Fill(DSClientes, "Clientes")
        DRClientes = DSClientes.Tables("Clientes").Rows(0)
    End Sub
    Function CargaFacturas() As Boolean
        DSFacturas = New DataSet
        StrSql = "select * from Documentos where estatus=1 and serie='WEB' and rfc='" + RFC + "' and year(fecha)=" + Str(Year(Now)) + " and month(fecha)<=" + Str(Month(Now))
        SDAFacturas = New SqlClient.SqlDataAdapter(StrSql, CargarTablas.conn)
        SDAFacturas.Fill(DSFacturas, "Facturas")
        If DSFacturas.Tables("Facturas").Rows.Count > 0 Then
            CargaFacturas = True
        Else
            CargaFacturas = False
        End If
    End Function

    Protected Sub BtnDescargarArchivo_Click(sender As Object, e As EventArgs) Handles BtnDescargarArchivo.Click

        Dim Indice As Integer = 0
        Dim Facts(), Fechas As String
        Dim fecha As New ListItem
        ViewState("Seleccionados") = LBxFacts.SelectedIndex > -1

        For Each fecha In LBxFacts.Items
            If fecha.Selected = True Then
                Facts(Indice) = "'" + Left(fecha.Text, 19) + "'"
                Indice += 1
            End If
        Next
        If Indice > 0 Then
            For i = 0 To Indice - 1
                If i > 0 Then
                    Fechas = Fechas + "," + Facts(i)
                Else
                    Fechas = Facts(i)
                End If
            Next
            StrSql = "Select archivoxml,archivopdf from Documentos where fecha in(" + Fechas + ")"
        Else
            'Response.Redirect("NoFacturas.aspx") 'poner mensaje
        End If
    End Sub

    Protected Sub LBxFacts_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LBxFacts.SelectedIndexChanged
        If LBxFacts.SelectedIndex > -1 Then


        End If
        'Response.Redirect("NoFacturas.aspx") 'poner mensaje
    End Sub
End Class