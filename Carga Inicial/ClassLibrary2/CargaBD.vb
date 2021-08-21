Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Web

Imports 

Public Class CargaBD

    Public DSEmpresa, DSDespachos, DSVtaProds, DSProductos, DSSeriefolio, DSClientes As DataSet
    Public SDAEmpresa, SDADespachos, SDAVtaProds, SDAProductos, SDASeriefolio, SDAClientes As SqlDataAdapter
    Public DREmpresa, DRDespachos, DRVtaProds, DRProductos, DRSeriefolio, DRClientes As DataRow
    Public QueryCarga, ResultF, RespBox As String

    Public Sub CargaCliente(XRFC As String)
        QueryCarga = "select * from EmpresasF "
        DSEmpresa = New DataSet
        SDAEmpresa = New SqlClient.SqlDataAdapter(QueryCarga, conn)
        SDAEmpresa.Fill(DSEmpresa, "EmpresasF")
        DREmpresa = DSEmpresa.Tables("EmpresasF").Rows(0)
    End Sub

End Class
