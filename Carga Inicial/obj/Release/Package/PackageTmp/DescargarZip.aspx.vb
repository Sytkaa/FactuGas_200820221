Imports System
Imports System.IO
Imports System.IO.Compression
Imports ClassLibrary1


Public Class DescargarZip
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim UserWeb As Integer = Val(Request.QueryString("UserWeb"))
        'SE ABRE PARA DESCARGA DEL ZIP
        Dim filePath As String = Request.QueryString("Archivozip")
        Dim Archivo As New FileInfo(filePath)
        If Archivo.Exists Then
            Response.Clear()
            Response.ClearHeaders()
            Response.ClearContent()
            Response.AddHeader("Content-Disposition", "attachment; filename=" + Archivo.Name)
            Response.AddHeader("Content-Length", Archivo.Length.ToString())
            Response.ContentType = "text/plain"
            Response.Flush()
            Response.TransmitFile(Archivo.FullName)
            Response.End()
            File.Delete(GeneraFactura.DescargaZip(UserWeb))
        Else
            Response.Redirect("Default.aspx")
        End If
    End Sub

End Class