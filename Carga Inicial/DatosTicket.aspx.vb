Imports System.Web.Hosting
Imports ClassLibrary1
Imports System.Math
Imports System
Imports System.IO
Imports ClassLibrary1.CalculaTotales
Imports System.Data.SqlClient

Public Class DatosTicket
    Inherits System.Web.UI.Page

    Dim DSDespachos, DSVtaProds, DSProductos, DSClientes, DSTempoFact As DataSet
    Dim SDADespachos, SDAVtaProds, SDAProductos, SDAClientes, SDATempoFact As SqlDataAdapter
    Dim DRDespachos, DRVtaProds, DRProductos, DRClientes, DRTempoFact As DataRow
    Dim Ticket, Importe, CadenaItem, ArchPdf, ArchXml, Mensaje, StrSql, ClaveCliente, RFC, NvaFactemp As String
    Dim UserWeb As Integer
    Dim Subtotal, IVA, IEPS, VolumenConcepto, ImporteBase, PrecioBase As Double
    Dim XTimbrado, BanderaOk As Boolean

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ClaveCliTxt = Request.QueryString("ClaveCliente")
        Dim RFCTxt = Request.QueryString("RFC")
        UserWeb = Request.QueryString("UserWeb")
        ClaveCliente = ClaveCliTxt
        RFC = RFCTxt
        LblNombre.Text = RFC
        LblInstancia.Text = CargarTablas.Instancia
        LblEspera.Visible = False
        BtnOk.Visible = False
        If CargarTablas.DREmpresa("Xtimbrado") Then
            LblTimbrado.Text = "Servidor SAT"
        Else
            LblTimbrado.Text = "Servidor de PRUEBAS!!"
        End If
        If GeneraFactura.CuentaLoads(UserWeb) = 1 Then
            LblEspera.Visible = True
            BtnGeneraFactura.Enabled = False
        End If
    End Sub

    Protected Sub BtnAgregar_Click(sender As Object, e As EventArgs) Handles BtnAgregar.Click
        Dim Ticket, Importe As String
        If Len(Trim(TxtTicket.Text)) < 4 Or Len(Trim(TxtTransacción.Text)) < 4 Then
            LblEspera.Text = "El formato del Folio del ticket o Transaccion es incorrecto!!!"
            MsgVisible()
        Else
            If VerificaFolio() Then
                If Left(Trim(TxtTicket.Text), 1).ToUpper <> "A" Then
                    If CargaTicket(TxtTransacción.Text, TxtTicket.Text) Then
                        Ticket = Trim(TxtTicket.Text)
                        Importe = Str(Round(DRDespachos("ImporteVenta"), 2))
                        CargaCombustible(Trim(DRDespachos("ClaveProd")))
                        VolumenConcepto = DRDespachos("Volumen")
                        CargaVenta("Combustible", DRDespachos("Volumen"), Trim(DRDespachos("ClaveProd")), DRDespachos("precio"), Trim(DRDespachos("Producto")), DRDespachos("importeventa"), DRProductos("ClaveSAT"), "LTR")
                    Else
                        LblEspera.Text = Mensaje
                        LblEspera.Visible = True
                        BtnOk.Visible = True
                    End If
                    TxtTicket.Text = ""
                    TxtTransacción.Text = ""
                Else
                    'INSERTAR AQUÍ VENTAPRODUCTOS
                    If CargaProductos(TxtTransacción.Text, TxtTicket.Text) Then
                        CargaVenta("Productos", DRProductos("Cantidad"), DRProductos("Codigo"), DRProductos("Precio"), DRProductos("Nombre"), DRProductos("Importe"), DRProductos("ClaveSAT"), DRProductos("CveUnidadSAT"))
                    End If
                End If
            End If
        End If
    End Sub

    Sub CargaVenta(Tipo As String, Cantidad As String, ClaveProd As String, Precio As String, Producto As String, Importe As String, ClaveSat As String, ClaveUnidad As String)
        Dim Fecha As String
        If Tipo = "Combustible" Then
            CalculaTotales.CalculaValores(Cantidad, Precio, DRProductos("nprodieps"), DRProductos("nprodiva"), Importe)
            IVA = CalculaTotales.CalcIVA
            Subtotal = CalculaTotales.CalcSubt
            ImporteBase = CalculaTotales.CalcImporteBase
            PrecioBase = CalculaTotales.CalcPrecioBase
            IEPS = CalculaTotales.CalcIEPS
            Ticket = Trim(TxtTicket.Text)
            Fecha = Format(DRDespachos("fechahora"), "dd/MM/yyyy hh:mm:ss")
        Else
            IVA = Round(Importe / 1.16 * 0.16, 2)
            Subtotal = Importe - IVA
            ImporteBase = Subtotal
            PrecioBase = Subtotal / Cantidad
            IEPS = 0
            Ticket = Mid(TxtTicket.Text, 2, Len(TxtTicket.Text) - 1)
            ClaveProd = Right(ClaveProd, 8)
            Fecha = Format(DRProductos("fecha"), "dd/MM/yyyy hh:mm:ss")
        End If
        CadenaItem = Trim(Ticket) + "|" + Producto + "|" + Trim(Importe)
        LBxTickets.Items.Add(CadenaItem)
        LblTotal0.Text = Val(LblTotal0.Text) + Val(Importe)
        LblTotal.Text = FormatCurrency(Val(LblTotal0.Text))
        NvaFactemp = "sp_nvatempfact33 " + ClaveCliente + ",0,'" + Fecha + "',2, " + Ticket + "," + Ticket + "," + Trim(ClaveProd) + ",'" + Trim(Producto) + "'," + Cantidad + _
         "," + Precio + "," + Str(ImporteBase) + "," + Str(Subtotal) + "," + Str(IVA) + "," + Str(PrecioBase) + _
          "," + Str(IEPS) + "," + Importe + ",'" + ClaveSat + "','" + ClaveUnidad + "'" + "," + Trim(TxtTransacción.Text)
        ComandoSQL.ComSQL(NvaFactemp)
        CargaTempofact()
        BtnGeneraFactura.Enabled = True
    End Sub

    Function VerificaFolio() As Boolean
        If Left(Trim(TxtTicket.Text), 1).ToUpper = "A" Then
            Ticket = Mid(TxtTicket.Text, 2, Len(TxtTicket.Text) - 1)
        Else
            Ticket = Trim(TxtTicket.Text)
        End If

        StrSql = "Select * from tempofact where FolioT=" + Ticket
        DSTempoFact = New DataSet
        SDATempoFact = New SqlClient.SqlDataAdapter(StrSql, CargarTablas.conn)
        SDATempoFact.Fill(DSTempoFact, "TempoFact")
        If DSTempoFact.Tables("TempoFact").Rows.Count > 0 Then
            VerificaFolio = False
            LblEspera.Text = "ERROR!! Ticket ya ingresado, sólo podrán agregarse una vez!!"
            MsgVisible()
        Else
            VerificaFolio = True
        End If
    End Function
    Protected Sub BtnGeneraFactura_Click(sender As Object, e As EventArgs) Handles BtnGeneraFactura.Click
        If GeneraFactura.CuentaLoads(UserWeb) = 0 Then
            GeneraFactura.CuentaLoads(UserWeb) = 1
            Dim Usocfdi() As String = Split(DDLUsoCfdi.Text, "|")
            Dim FormaPago() As String = Split(DDLFormaPago.Text, "|")
            Response.Redirect("GenerandoFactura.aspx?Clavecliente=" + ClaveCliente + "&UsoCfdi=" + Usocfdi(0) + "&FormaPago=" + FormaPago(0) + "&UserWeb=" + Str(UserWeb), False)
        Else
            Response.Redirect("Default.aspx")
        End If

    End Sub

    Protected Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOk.Click
        If BanderaOk Then
            Response.Redirect("Default.aspx")
        Else
            TxtTicket.Text = ""
            TxtTransacción.Text = ""
            LblEspera.Visible = False
            BtnOk.Visible = False
            BtnAgregar.Enabled = True
        End If

    End Sub

    Public Sub MsgVisible()
        LblEspera.Visible = True
        BtnOk.Visible = True
    End Sub

    Function CargaTicket(Transacc As String, Numtick As String) As Boolean
        DSDespachos = New DataSet
        SDADespachos = New SqlClient.SqlDataAdapter("select * from Despachos where Transaccion=" + Transacc + " and Numticket=" + Numtick, CargarTablas.conn)
        SDADespachos.Fill(DSDespachos, "Despachos")
        If DSDespachos.Tables("Despachos").Rows.Count > 0 Then
            DRDespachos = DSDespachos.Tables("Despachos").Rows(0)
            CargaTicket = False
            If DRDespachos("Factura") > 0 And (DRDespachos("EstadoFact") = 1 Or DRDespachos("EstadoFact") = 3) Then
                Mensaje = "El Ticket solicitado ya fué facturado anteriormente"
            ElseIf DRDespachos("Solicitud") > 0 Then
                Mensaje = "El Ticket solicitado es de Crédito, Imposible Facturar"
            ElseIf Month(DRDespachos("Fechahora")) < Month(Now) Then
                Mensaje = "Tickets de meses anteriores no se pueden facturar vía WEB"
            Else
                CargaTicket = True
            End If
        Else
            Mensaje = "Folio de Ticket No Encontrado!!"
            CargaTicket = False
        End If
    End Function

    Function CargaProductos(Transacc As String, Numtick As String) As Boolean
        
        Numtick = Mid(Numtick, 2, Len(Numtick) - 1)
        DSProductos = New DataSet
        StrSql = "SELECT ventaproductos.codigo, ventaproductos.nombre, cantidad, ventaproductos.precio, importe, numeroventa, Fecha, ID, ISNULL(factura, 0) AS factura,ClaveSAT,CveUnidadSat " + _
            "FROM ventaproductos inner join ProductosV on ventaproductos.codigo=Productosv.Codigo where Id=" + Transacc + " and Numeroventa=" + Numtick
        SDAProductos = New SqlClient.SqlDataAdapter(StrSql, CargarTablas.conn)
        SDAProductos.Fill(DSProductos, "VentaProductos")
        If DSProductos.Tables("VentaProductos").Rows.Count > 0 Then
            
            DRProductos = DSProductos.Tables("VentaProductos").Rows(0)
            CargaProductos = False
            If DRProductos("Factura") > 0 Then
                Mensaje = "El Ticket solicitado ya fué facturado anteriormente"
            ElseIf Month(DRProductos("Fecha")) <> Month(Now) Then
                Mensaje = "Tickets de meses anteriores no se pueden facturar vía WEB"
            Else
                CargaProductos = True
            End If
        Else
            Mensaje = "Ticket No Encontrado!!"
            CargaProductos = False
        End If
    End Function

    Sub CargaCombustible(ClaveProducto As String)
        DSProductos = New DataSet
        SDAProductos = New SqlClient.SqlDataAdapter("Select * from productos where NProdClave='" + ClaveProducto + "'", CargarTablas.conn)
        SDAProductos.Fill(DSProductos, "productossF")
        DRProductos = DSProductos.Tables("productossF").Rows(0)
    End Sub

    Sub CargaTempofact()
        StrSql = "Select cliente,claveprod,unidades,nombre,precio,importe,subtotal,iva,importeb,tipot,foliot,preciob,unidadmedida,ClaveSAT,CveUnidadSat,ieps,Fecha,Transaccion" + _
            " from tempofact where cliente=" + ClaveCliente
        DSTempoFact = New DataSet
        SDATempoFact = New SqlClient.SqlDataAdapter(StrSql, CargarTablas.conn)
        SDATempoFact.Fill(DSTempoFact, "TempoFact")
    End Sub

End Class