Imports System.Web.Hosting
Imports ClassLibrary1
Imports System.Math
Imports System
Imports System.IO
Imports System.IO.Compression
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Mail
Imports FactXMLdll.XmlFact
Imports System.Web.UI

Public Class GenerandoFactura
    Inherits System.Web.UI.Page

    Dim DSClientes, DSTempoFact, DSFolioFactsWEB As DataSet
    Dim SDAClientes, SDATempoFact, SDAFolioFactsWEB As SqlDataAdapter
    Dim DRClientes, DRTempoFact, DRFolioFactsWEB As DataRow
    Dim SubT, IvaT, IepsT, TotalT, CantV, IVA, Total As Double
    Dim XTIMBRADO, MostrarPDF, Reenvio, BanderaREP, BanderaReLoadPage As Boolean
    Dim objJSON As Object
    Dim oScriptEngine As MSScriptControl.ScriptControl
    Dim TasaOCuota, CadenaOriginal, Logo, trsubt, striva, strtotal, TipoFac, FechAct, instancia, strsql, CtaPred, MetPago, FechaId, sJsonString, XToken, UrlT, MES(11) As String
    Public PCnombre, Mensaje, Etiqueta, Serie, FechaInicial, FechaFinal, Vigencia, EquipoVpn(), TxtXML, UUID, XmlData, DirectorioFrom, DirectorioTo As String
    Public FolioFactura, ClaveCliente, Counter, CounterLog, UserWeb As Integer
    Dim Nombre, Calle, NumExterior, NumInterior, Colonia, Localidad, Municipio, Estado, Pais, CP, RFCCliente, ArchivoXmlFrom, ArchivoPdfFrom, ArchivoZip As String
    Dim AdmingasPath, ArCer, ArKey, ArchivoXml, Archivopdf, Ress, TxtResponse, email, Xusocfdi, Xforpago, Certificado, NoCertificado, TextLog As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ClaveCliTxt = Request.QueryString("ClaveCliente")
        Dim UsoCfdiTxt = Request.QueryString("UsoCfdi")
        Dim FormaPagoTxt = Request.QueryString("FormaPago")
        UserWeb = Request.QueryString("UserWeb")
        If GeneraFactura.CuentaLoads(UserWeb) = 1 Then
            GeneraFactura.CuentaLoads(UserWeb) = 2
            ClaveCliente = Val(ClaveCliTxt)
            Xusocfdi = UsoCfdiTxt
            Xforpago = FormaPagoTxt
            XLog("Entró a load Generando Factura")
            CargaCliente()
            If ProcesaCFDI() Then
                GeneraFactura.CuentaLoads(UserWeb) = 3
                LblMensaje.Text = Mensaje
                BtnDescargarArchivo.Visible = True
            Else

            End If
            else
        End If
        LblLog.Text = "Cuentaloads=" + Str(GeneraFactura.CuentaLoads(UserWeb)) + " Userweb=" + Str(UserWeb)
        TxtLog.Text = Mensaje
    End Sub

    Public Function ProcesaCFDI() As Boolean
        UUID = ""
        ProcesaCFDI = False
        XLog("Comienza Generar " + Mensaje)
        Dim SubT, IvaT, TotalT As Double
        Dim Instancia() As String
        XLog("Antes de Generar CFDI " + Mensaje)
        Try
            If GeneraCFDI(Xusocfdi, Xforpago, DRClientes("RFC"), DRClientes("Nombre")) Then 'GENERA EL XML TIMBRADO
                XLog("Despues de Generar CFDI " + Mensaje)
                If CargaTempofact() Then
                    For i = 0 To DSTempoFact.Tables("TempoFact").Rows.Count - 1
                        DRTempoFact = DSTempoFact.Tables("TempoFact").Rows(i)
                        SubT += DRTempoFact("Subtotal")
                        IvaT += DRTempoFact("Iva")
                        TotalT += DRTempoFact("Importe")
                    Next
                    Instancia = Split(CargarTablas.Instancia, "\")
                    Dim UbicacionXml, UbicacionPdf As String
                    UbicacionXml = Replace(ArchivoXml, "C:\ArchivosWEB", "\")
                    UbicacionXml = Replace(UbicacionXml, "ArchivosXML", "Users\Public\Admingas\ArchivosXML")
                    UbicacionPdf = Replace(Archivopdf, "C:\ArchivosWEB", "\")
                    UbicacionPdf = Replace(UbicacionPdf, "ArchivosXML", "Users\Public\Admingas\ArchivosXML")
                    FechAct = FactXMLdll.XmlFact.FechaFactura
                    strsql = "sp_nvafact 'WEB'," + Str(FolioFactura) + ",'Cliente WEB',0,0,'" + FechAct + "',1," + Str(IvaT) + "," + Str(CargarTablas.DREmpresa("IVA")) + _
                        "," + Str(SubT) + "," + Str(TotalT) + ",'" + Trim(DRClientes("Nombre")) + "'," + Str(DRClientes("Clave")) + ",'" + DRClientes("RFC") + "','" + Trim(DRClientes("Calle")) + _
                        "','" + Trim(DRClientes("NumExterior")) + "','" + Trim(DRClientes("NumInterior")) + "','" + Trim(DRClientes("Colonia")) + "','" + Trim(DRClientes("Localidad")) + "','" + Trim(DRClientes("Municipio")) + "','" + _
                    Trim(DRClientes("Estado")) + "','" + Trim(DRClientes("Pais")) + "','" + DRClientes("CP") + "','" + Trim(DRClientes("email")) + "','PUE|" + Xforpago + "',0,0,0,'" + UbicacionXml + "','" + UbicacionPdf + "',0,' '"
                    XLog("Antes de ejecutar sp_nvafact " + strsql)
                    ComandoSQL.ComSQL(strsql) 'PROBAR PRIMERO DIRECTO EN SQL MANAGER
                    XLog("despues de sp_nvafact " + Mensaje)
                    strsql = "Update Documentos set UUID='" + UUID + "' where folioint=" + Str(FolioFactura)
                    ComandoSQL.ComSQL(strsql)
                    XLog("despues Update Documentos")
                    For i = 0 To DSTempoFact.Tables("TempoFact").Rows.Count - 1
                        DRTempoFact = DSTempoFact.Tables("TempoFact").Rows(i)
                        strsql = "sp_MOVFACT33 " + Str(FolioFactura) + ",'" + Trim(DRTempoFact("ClaveProd")) + "','" + Str(DRTempoFact("Unidades")) + "',LITROS,'" + _
                        Trim(DRTempoFact("Nombre")) + "'," + Str(DRTempoFact("Precio")) + "," + Str(DRTempoFact("SubTotal")) + "," + Str(DRTempoFact("PrecioB")) + _
                        "," + Str(DRTempoFact("IVA")) + "," + Str(DRTempoFact("Importe")) + "," + Str(DRTempoFact("FolioT")) + ",'" + _
                        Format(DRTempoFact("Fecha"), "dd/MM/yyyy hh:mm:ss") + "'," + Str(Month(DRTempoFact("Fecha"))) + ",1,'" + Str(DRTempoFact("ClaveSAT")) + "','" + _
                        Str(DRTempoFact("ClaveSAT")) + "'," + Str(DRTempoFact("IEPS")) + ",'WEB'"
                        XLog("Antes de ejecutar sp_MOVFACT33 " + strsql)
                        ComandoSQL.ComSQL(strsql)
                        XLog("despues de sp_MOVFACT33 ")
                    Next
                    XLog("Antes de ejecutar sp_tfotodos " + strsql)
                    ComandoSQL.ComSQL("sp_tfotodos " + Str(FolioFactura) + ", " + Str(DRClientes("Clave")))
                    XLog("Despues de ejecutar sp_tfotodos ")
                    CargarTablas.MensajeMail = "Estimado cliente:" & vbNewLine & "Anexo encontrará el comprobante CFDI de la factura: WEB" + Str(FolioFactura) + " en formato PDF y XML" & vbNewLine & _
                        "Agradecemos su preferencia." & vbNewLine & "Este es un mensaje generado automáticamente por el sistema." & vbNewLine & _
                        "Favor de no responder al mismo."
                    CargarTablas.FolioEnvio = Str(FolioFactura)
                    XLog("Antes de Enviar email " + Mensaje)
                    EnviarCFDI(Trim(DRClientes("Email")))
                    XLog("Despues de Enviar email " + Mensaje)
                   
                    'SE COMPRIME ARCHIVO XML->ZIP Y SE COLOCA EN DIRECTORIO POR ENVIAR

                    DirectorioFrom = "C:\ArchivosWEB\DirectorioFrom"
                    DirectorioTo = "C:\ArchivosWEB\DirectorioTo"
                    ArchivoXmlFrom = Replace(ArchivoXml, "C:\ArchivosWEB\" + Trim(CargarTablas.DREmpresa("Nombre_CC")) + "\ArchivosXML\", "C:\ArchivosWEB\DirectorioFrom\")
                    ArchivoPdfFrom = Replace(Archivopdf, "C:\ArchivosWEB\" + Trim(CargarTablas.DREmpresa("Nombre_CC")) + "\ArchivosXML\", "C:\ArchivosWEB\DirectorioFrom\")
                    If Directory.Exists(DirectorioFrom) = False Then
                        Directory.CreateDirectory(DirectorioFrom)
                    End If
                    If Directory.Exists(DirectorioTo) = False Then
                        Directory.CreateDirectory(DirectorioTo)
                    End If
                    FileCopy(ArchivoXml, ArchivoXmlFrom)
                    FileCopy(Archivopdf, ArchivoPdfFrom)
                    ArchivoZip = Replace(ArchivoXmlFrom, "xml", "zip")
                    ArchivoZip = Replace(ArchivoZip, "From", "To")
                    ZipFile.CreateFromDirectory(DirectorioFrom, ArchivoZip)
                    File.Delete(ArchivoXmlFrom)
                    File.Delete(ArchivoPdfFrom)
                    GeneraFactura.DescargaZip(UserWeb) = ArchivoZip
                    Dim XmlUpload, PdfUpload As String
                    XmlUpload = ""
                    PdfUpload = ""
                    XmlUpload = Replace(ArchivoXml, AdmingasPath + "\" + Trim(CargarTablas.DREmpresa("Nombre_CC")) + "\ArchivosXML\", "/")
                    XmlUpload = "ftp://" + Instancia(0) + XmlUpload
                    PdfUpload = Replace(Archivopdf, AdmingasPath + "\" + Trim(CargarTablas.DREmpresa("Nombre_CC")) + "\ArchivosXML\", "/")
                    PdfUpload = "ftp://" + Instancia(0) + PdfUpload
                    XLog("Antes de subir los archivos pdf y xml a la CC (por ftp)" + Mensaje)
                    'SE SUBEN AL CC CORRESPONDIENTE VIA FTP
                    My.Computer.Network.UploadFile(Archivopdf, PdfUpload)
                    My.Computer.Network.UploadFile(ArchivoXml, XmlUpload)
                    XLog("Termina Funcion" + Mensaje)
                    LblMensaje.Visible = True
                    BtnCerrar.Visible = True
                Else
                    Mensaje = "ERROR: No se Localizaron Tickets en TempoFact, imposible facturar"
                End If
                BtnCerrar.Visible = True
                ProcesaCFDI = True
            Else
                XLog("NO SE PUDO GENERAR LA FACTURA!! ERROR,: " + Mensaje)
                Mensaje = "NO SE PUDO GENERAR LA FACTURA!! ERROR,: " + Mensaje
                ProcesaCFDI = False
            End If
            LblMensaje.Text = "UserWeb: " + Str(UserWeb) + "  FOLIO: " + Str(FolioFactura) + "RFC: " + DRClientes("RFC") + Mensaje
        Catch ex As Exception
            'Mensaje = ex.Message
            XLog("ERROR: " + ex.Message)
            XLog("Después de llamar a descargas")
            Response.Redirect("DescargarZip.aspx?Archivozip=" + ArchivoZip)

        End Try

    End Function

    Protected Sub BtnCerrar_Click(sender As Object, e As EventArgs) Handles BtnCerrar.Click
        File.Delete(GeneraFactura.DescargaZip(UserWeb))
        Response.Redirect("Default.aspx")
    End Sub

    Function GeneraCFDI(UsoCfdi As String, ForPago As String, RFC As String, Rsocial As String) As Boolean
        '/// GENERAR FACTURA ///
        XLog("Entró a funciónGeneraCFDI " + Mensaje)
        EquipoVpn = Split(CargarTablas.Instancia, "\")
        AdmingasPath = "C:\ArchivosWEB"
        ArCer = AdmingasPath + CargarTablas.DREmpresa("ArchivoCer") 'Verificar
        ArKey = AdmingasPath + CargarTablas.DREmpresa("ArchivoKey") 'Verificar
        FolioFactura = FolioNvo() 'Verificar
        XLog("Antes de Generar FolioFactura " + Mensaje)
        If FolioFactura = 0 Then
            Mensaje = "ERROR: No se pudo generar el folio de la Factura WEB, intente de nuevo, Gracias"
            GeneraCFDI = False
            Exit Function
        End If
        XLog("Después de Generar Folio " + Mensaje)
        XLog("Antes de Crear Xml " + Mensaje)
        TxtResponse = FactXMLdll.XmlFact.CreaTxtXML(ClaveCliente, "WEB", FolioFactura, UsoCfdi, ForPago, RFC, Rsocial, CargarTablas.Instancia, AdmingasPath)
        XLog("Antes de Timbrar " + TxtResponse)
        Ress = FactXMLdll.XmlFact.Timbrar(CargarTablas.DREmpresa("Sender"), CargarTablas.DREmpresa("paswSender"), CargarTablas.DREmpresa("XTimbrado"), TxtResponse)
        UUID = FactXMLdll.XmlFact.UUID
        XmlData = FactXMLdll.XmlFact.XmlData
        XLog("Validar: " + Ress)
        If Ress = "success" Then 'VALIDA TIMBRADO EXITOSO
            ArchivoXml = AdmingasPath + "\" + Trim(CargarTablas.DREmpresa("Nombre_CC")) + "\ArchivosXML\" + Trim(CargarTablas.DREmpresa("rfc")) + "_WEB" + Trim(Format(FolioFactura, "00#")) + ".xml"
            Archivopdf = AdmingasPath + "\" + Trim(CargarTablas.DREmpresa("Nombre_CC")) + "\ArchivosXML\" + Trim(CargarTablas.DREmpresa("rfc")) + "_WEB" + Trim(Format(FolioFactura, "00#")) + ".pdf"
            XLog("Antes de escribir el archivo xml timbrado " + Mensaje)
            Dim escritor As New StreamWriter(ArchivoXml)
            escritor.WriteLine(Trim(XmlData))
            escritor.Close()
            Dim ApPDF As String
            Dim Xusuario As String = "U/Cliente WEB"
            ApPDF = """" + AdmingasPath + "\XPN.EXE""" + " " + """" + ArchivoXml + """" + " """ + AdmingasPath + "\" + CargarTablas.DREmpresa("Logo") + """" + " """ + Xusuario + ""
            Call Shell(ApPDF, AppWinStyle.Hide)
            GeneraCFDI = True
        Else
            XLog("ERROR:" + FactXMLdll.XmlFact.sJsonString)
            GeneraCFDI = False
        End If
        strsql = "delete from FoliofactsWEB where folio=" + Str(FolioFactura)
        ComandoSQL.ComSQL(strsql)

    End Function

   
    Function LeerArch(arch As String)
        Dim TextArch = "", LineText As String
        Using streamReader As System.IO.StreamReader = System.IO.File.OpenText(arch)
            While Not streamReader.EndOfStream
                LineText = streamReader.ReadLine()
                TextArch = TextArch & LineText
            End While
        End Using
        LeerArch = TextArch
    End Function

    Sub XLog(CadenaLog As String)
        Counter += 1
        TextLog = "C:\ArchivosWEB\" + CargarTablas.DREmpresa("Nombre_CC") + "\ArchivosXml\Log" + Str(UserWeb) + ".txt"
        If Counter = 1 Then
        Dim escritor As New StreamWriter(TextLog)
            escritor.WriteLine("Log " + Now.ToString)
            escritor.Close()
        Else
            Dim escritor As StreamWriter
            escritor = File.AppendText(TextLog)
            escritor.WriteLine(Trim(Str(Counter)) + " " + CadenaLog)
            escritor.Flush()
            escritor.Close()
        End If

    End Sub

    Function EncryptSHA256Managed(ByVal ClearString As String) As String
        Dim uEncode As New UnicodeEncoding()
        Dim bytClearString() As Byte = uEncode.GetBytes(ClearString)
        Dim sha As New  _
        System.Security.Cryptography.SHA256Managed()
        Dim hash() As Byte = sha.ComputeHash(bytClearString)
        Return Convert.ToBase64String(hash)
    End Function

    Function LeeCertificado(ByVal fileName As String) As Byte()
        Dim f As New FileStream(fileName, FileMode.Open, FileAccess.Read)
        Dim size As Integer = Fix(f.Length)
        Dim data(size) As Byte
        size = f.Read(data, 0, size)
        f.Close()
        Return data
    End Function

    Function ConvertFileToBase64(ByVal fileName As String) As String
        Return Convert.ToBase64String(System.IO.File.ReadAllBytes(fileName))
    End Function
   
    Sub CargaCliente()
        DSClientes = New DataSet
        SDAClientes = New SqlClient.SqlDataAdapter("select * from Clientes where clave='" + Str(ClaveCliente) + "'", CargarTablas.conn)
        SDAClientes.Fill(DSClientes, "Clientes")
        DRClientes = DSClientes.Tables("Clientes").Rows(0)
    End Sub

    Function CargaTempofact() As Boolean
        strsql = "Select cliente,claveprod,unidades,nombre,precio,importe,subtotal,iva,importeb,tipot,foliot,preciob,unidadmedida,ClaveSAT,CveUnidadSat,ieps,Fecha,Transaccion" + _
            " from tempofact where cliente=" + Str(ClaveCliente)
        DSTempoFact = New DataSet
        SDATempoFact = New SqlClient.SqlDataAdapter(strsql, CargarTablas.conn)
        SDATempoFact.Fill(DSTempoFact, "TempoFact")
        If DSTempoFact.Tables("TempoFact").Rows.Count > 0 Then
            CargaTempofact = True
        Else
            CargaTempofact = False
        End If
    End Function

    Function FolioNvo() As Integer
        strsql = "INSERT INTO FolioFactsWEB DEFAULT VALUES"
        ComandoSQL.ComSQL(strsql)
        DSFolioFactsWEB = New DataSet
        SDAFolioFactsWEB = New SqlClient.SqlDataAdapter("select Folio from FolioFactsWEB order by Folio desc", CargarTablas.conn)
        SDAFolioFactsWEB.Fill(DSFolioFactsWEB, "FolioFactsWEB")
        If DSFolioFactsWEB.Tables("FolioFactsWEB").Rows.Count > 0 Then
            DRFolioFactsWEB = DSFolioFactsWEB.Tables("FolioFactsWEB").Rows(0)
            FolioNvo = DRFolioFactsWEB("Folio")
        Else
            FolioNvo = 0
        End If
    End Function

    Sub EnviarCFDI(EmailCliente As String)
        EmailCliente = Replace(EmailCliente, ";", ",")
        Dim smtp As New SmtpClient
        Dim correo As New MailMessage()
        Dim adjunto As Attachment
        Dim date2 As Date = Now
        Dim tiempogen As Long
        Do While Not My.Computer.FileSystem.FileExists(Archivopdf)
            Dim date1 As Date = Now
            tiempogen = DateDiff(DateInterval.Second, date2, date1)
            If tiempogen > 15 Then
                Mensaje = "Factura WEB" + Str(FolioFactura) + " generada y timbrada exitosamente: NO SE PUDO GENERAR EL PDF!!, FAVOR DE BAJARLA DEL PORTAL DEL SAT"
                Exit Sub
            End If
        Loop
        tiempogen = 0
        date2 = Now
        Do While ArchivoEnUso(Archivopdf)
            Dim date1 As Date = Now
            tiempogen = DateDiff(DateInterval.Second, date2, date1)
            If tiempogen > 10 Then
                Mensaje = "No se recibió el PDF, favor de enviar el correo manualmente!!!"
                Exit Sub
            End If
        Loop
        tiempogen = 0
        date2 = Now
        Dim UsuarioSmtp As String = Trim(CargarTablas.DREmpresa("usere")), ClaveSmtp As String = Trim(CargarTablas.DREmpresa("paswe"))
        With smtp
            '.UseDefaultCredentials = False
            .Credentials = New Net.NetworkCredential(UsuarioSmtp, ClaveSmtp)
            .Port = 587
            .Host = "smtp.gmail.com"
            .EnableSsl = True
        End With
        correo = New MailMessage()
        With correo
            Dim Asunto As String
            Asunto = CargarTablas.DREmpresa("asunto") + " de " + CargarTablas.DREmpresa("de")
            .From = New MailAddress(UsuarioSmtp)
            .To.Add(Trim(EmailCliente))
            .Subject = Asunto
            .Body = CargarTablas.MensajeMail
            .IsBodyHtml = False
            .Priority = MailPriority.Normal
            adjunto = New Attachment(Archivopdf)
            .Attachments.Add(adjunto)
            adjunto = New Attachment(ArchivoXml)
            .Attachments.Add(adjunto)
        End With
        Try
            smtp.Send(correo)
            Mensaje = "Factura WEB" + Str(FolioFactura) + " generada y enviada al correo exitosamente!!"
        Catch ex As Exception
            Mensaje = "Factura WEB" + Str(FolioFactura) + " generada exitosamente    (No se pudo enviar al correo, favor de descargar.)"
        End Try
        'LblMensaje.Text = Mensaje
    End Sub

    Function ArchivoEnUso(filePath As String) As Boolean
        Dim rtnvalue As Boolean = False
        Try
            Dim fs As System.IO.FileStream = System.IO.File.OpenWrite(filePath)
            fs.Close()
        Catch ex As System.IO.IOException
            rtnvalue = True
        End Try
        Return rtnvalue
    End Function

    Protected Sub BtnDescargarArchivo0_Click(sender As Object, e As EventArgs) Handles BtnDescargarArchivo.Click
        'ELIMINAMOS ARCHIVOS Y DIRECTORIO TEMPORALES
        Response.Redirect("DescargarZip.aspx?Archivozip=" + GeneraFactura.DescargaZip(UserWeb) + "&UserWeb=" + Str(UserWeb))
    End Sub
   
End Class