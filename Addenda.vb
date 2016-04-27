
Public Class Addenda
    Public Property Emisor As Emisor
    Public Property Receptor As Receptor
    Public Property Documento As Documento
End Class
Public Class Emisor
    Public Property Web As String
    Public Property Telefono As String
    Public Property LargaDistancia As String
    Public Property Fax As String
End Class
Public Class Receptor
    Public Property Codigo As String
    Public Property RefDeposito As String
End Class
Public Class Documento
    Public Property OrdenCompra As String
    Public Property ImportePP As Decimal
    Public Property FechaPP As String
    Public Property FechaVence As String
    Public Property ImporteLetra As String
    Public Property TotalArticulos As Integer
    Public Property RefPedido As String
    Public Property Observaciones As List(Of Observacion)
    Public Property Agente As String
    Public Property Movimientos As List(Of Movimiento)
    Public Property TituloDocumento As String
End Class
Public Class Movimiento
    Public Property PorcDescto As Decimal
    Public Property Detalle As String
End Class

Public Class Observacion
    Public Property Detalle As String
End Class





