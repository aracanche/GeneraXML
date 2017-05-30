Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Net

Module Module1
    Private XMLOrigen As String
    Private XMLAddenda As String

    Private TipoDocumento As FormatoPDF    
    Private Folio As Integer
    Private DestinoPDF As String
    Private cAddenda As New Addenda
    Private bEsCopia As Boolean
    Private bEsSaldoInicial As Boolean

    Public Empresa As String
    Public UrlGeneraPDF As String
    Public RutaXMLTemp As String
    Public RutaPDFTemp As String

    Private sqlConnSAP As SqlConnection

    Private bEsParametro As Boolean

    Public Enum FormatoPDF
        Devolucion
        Factura
        FacturaIVA
        FacturaRenta
        NotaCredito
        PagoParcial
        NotaCargo
    End Enum

    Sub Main()
        Try
            'Leer el archivo Facturas.TXT que debe estar en la misma ubicación que el EXE
            Dim DirCurr As String = System.AppDomain.CurrentDomain.BaseDirectory
            If Not DirCurr.EndsWith("\") Then
                DirCurr += "\"
            End If
            Dim sFile As String = ""
            Dim sCommand() As String = Environment.GetCommandLineArgs()
            If sCommand.Length = 1 Then
                sFile = DirCurr + "Facturas.TXT"
                bEsParametro = False
            Else
                sFile = sCommand(1)
                bEsParametro = True
            End If
            If Not File.Exists(sFile) Then
                Exit Sub
            End If
            Dim objReader As New StreamReader(sFile)
            Dim sLine As String = ""
            Dim Parametros() As String
            Dim Param As String
            Dim iParam As Short
            Dim sEmpresa As String
            Do
                Try
                    sLine = objReader.ReadLine()
                    If Not sLine Is Nothing Then
                        Parametros = sLine.Split("*")
                        For iParam = 1 To Parametros.Length - 1
                            Param = Parametros(iParam)
                            Select Case Param.Substring(0, 2)
                                Case "A-" 'Ruta o archivo XML
                                    XMLOrigen = Param.Replace("A-", "").Trim
                                Case "B-" 'Tipo de documento
                                    TipoDocumento = Val(Param.Replace("B-", "").Trim)
                                Case "C-" 'Empresa
                                    sEmpresa = Param.Replace("C-", "").Trim
                                    If sEmpresa <> Empresa Then
                                        If Not IsNothing(Empresa) Then
                                            DesconectaSAPSQL()
                                            DesconectaSQL()
                                            Threading.Thread.Sleep(500)
                                        End If
                                        Empresa = sEmpresa
                                    End If
                                Case "D-" 'Folio
                                    Folio = Val(Param.Replace("D-", "").Trim)
                                Case "E-" 'Destino
                                    DestinoPDF = Param.Replace("E-", "").Trim
                                Case "F-" 'Original o copia
                                    bEsCopia = IIf(Val(Param.Replace("F-", "").Trim) = 1, True, False)
                            End Select
                        Next
                        If XMLOrigen.ToLower.EndsWith(".xml") Then
                            bEsSaldoInicial = True
                        Else
                            bEsSaldoInicial = False
                            If Not XMLOrigen.EndsWith("\") Then
                                XMLOrigen += "\"
                            End If
                        End If

                        If Not DestinoPDF.EndsWith("\") Then
                            DestinoPDF += "\"
                        End If

                        Dim LlenoDatosAddenda As Boolean = False
                        If bEsSaldoInicial Then
                            If File.Exists(XMLOrigen) Then
                                LlenoDatosAddenda = GeneraAddendaCompac()
                            End If
                        Else
                            If Directory.Exists(XMLOrigen) Then
                                LlenoDatosAddenda = GeneraAddendaSAP()

                            End If
                        End If
                        If LlenoDatosAddenda Then
                            If Directory.Exists(DestinoPDF) Then
                                GeneraPDF()
                            End If
                        End If
                    End If
                Catch ex As Exception
                    'Agregar a un log el error ocasionado
                End Try
            Loop Until sLine Is Nothing
            objReader.Close()

            'Cerramos las conexiones
            DesconectaSAPSQL()
            DesconectaSQL()
        Catch ex As Exception
            If bEsParametro Then
                MsgBox("Error Main: " + ex.Message.ToString, vbExclamation)
            End If
        End Try


    End Sub

    Private Function AbreConexionSAP() As Boolean
        Try
            Dim CadenaConexion As String = "Data Source=FareTDB1\sap;Initial Catalog=" + Empresa.ToString + ";Persist Security Info=True;User ID=sa;Password=B1Admin"
            sqlConnSAP = New SqlConnection(CadenaConexion)
            sqlConnSAP.Open()
            If sqlConnSAP.State = ConnectionState.Open Then
                SetSQLSAP("Set dateFormat 'ymd'")
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            If bEsParametro Then
                MsgBox("Error AbreConexionSAP: " + ex.Message.ToString, vbExclamation)
            End If
            Return False
        End Try

    End Function

    Public Function GetSQLSAP(ByVal sSql As String) As DataTable
        Dim dT As New DataTable
        If IsNothing(sqlConnSAP) Then
            If Not AbreConexionSAP Then
                Return dT
            End If
        End If
        Dim sqlCmd As New SqlCommand(sSql, sqlConnSAP)
        Dim sqlDA As SqlDataAdapter
        sqlDA = New SqlDataAdapter(sSql, sqlConnSAP)
        sqlDA.Fill(dT)
        Return dT
    End Function

    Public Sub DesconectaSAPSQL()
        Try
            sqlConnSAP.Close()
            sqlConnSAP = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Public Function SetSQLSAP(ByVal sSql As String) As Integer
        If IsNothing(sqlConnSAP) Then
            If Not AbreConexionSAP() Then
                Return 0
            End If
        End If
        Dim sqlCmd As New SqlCommand(sSql, sqlConnSAP)
        Dim RegistrosAfectados As Integer
        RegistrosAfectados = sqlCmd.ExecuteNonQuery()
        Return RegistrosAfectados
    End Function

    Private Function GeneraAddendaSAP() As Boolean
        Try
            cAddenda = New Addenda
            Dim cEmisor As New Emisor
            Dim cReceptor As New Receptor
            Dim cDocumento As New Documento
            Dim cMovi As Movimiento

            Dim UUID As String = ""
            Dim FechaDocumento As Date

            cDocumento.Movimientos = New List(Of Movimiento)

            Dim sSql As String
            Dim dT As DataTable
            'Emisor
            sSql = "Select isnull(IntrntAdrs,'')IntrntAdrs from ADM1"
            dT = GetSQLSAP(sSql)
            If dT.Rows.Count > 0 Then
                cEmisor.Web = dT.Rows(0).Item("IntrntAdrs")
            End If
            sSql = "Select isnull(Phone1,'')Phone1, isnull(Phone2,'')Phone2, isnull(Fax,'')Fax from OADM"
            dT = GetSQLSAP(sSql)
            If dT.Rows.Count > 0 Then
                cEmisor.Telefono = dT.Rows(0).Item("Phone1")
                cEmisor.LargaDistancia = dT.Rows(0).Item("Phone2")
                cEmisor.Fax = dT.Rows(0).Item("Fax")
            End If

            'Documento
            Select Case TipoDocumento
                Case FormatoPDF.Factura, FormatoPDF.FacturaIVA, FormatoPDF.FacturaRenta
                    sSql = "Select CardCode, isnull(DiscPrcnt,0)DiscPrcnt, SlpCode, U_OrdenCompra, U_TotalPP, U_FechaPP, DocDueDate, Comments, DocEntry, DocTotal, DocDate, EDocNum, u_Metodo from OINV where DocSubType='--' and DocNum=" + Folio.ToString
                Case FormatoPDF.Devolucion, FormatoPDF.NotaCredito
                    sSql = "Select CardCode, isnull(DiscPrcnt,0)DiscPrcnt, SlpCode, U_OrdenCompra, U_TotalPP, U_FechaPP, DocDueDate, Comments, DocEntry, DocTotal, DocDate, EDocNum, u_Metodo from ORIN where DocNum=" + Folio.ToString
                Case FormatoPDF.NotaCargo
                    sSql = "Select CardCode, isnull(DiscPrcnt,0)DiscPrcnt, SlpCode, U_OrdenCompra, U_TotalPP, U_FechaPP, DocDueDate, Comments, DocEntry, DocTotal, DocDate, EDocNum, u_Metodo from OINV where  DocSubType='dn' and DocNum=" + Folio.ToString
            End Select
            'sSql += " and EDocNum is Not null"
            dT = GetSQLSAP(sSql)
            Dim DocEntry As Integer
            Dim TotalDocumento As Decimal
            Dim PorcDescDoc As Decimal
            Dim Fecha As Date
            Dim Agente As Integer
            Dim MetPago As String

            Dim dR As DataRow

            If dT.Rows.Count > 0 Then
                dR = dT.Rows(0)
                cReceptor.Codigo = dR("CardCode")
                Fecha = IIf(IsDate(dR("U_FechaPP")), dR("U_FechaPP"), dR("DocDueDate"))
                cDocumento.FechaPP = Fecha.ToString("dd 'de' MMMM 'de' yyyy")
                Fecha = dR("DocDueDate")
                cDocumento.FechaVence = Fecha.ToString("dd 'de' MMMM 'de' yyyy")
                cDocumento.ImportePP = Math.Round(IIf(IsDBNull(dR("U_TotalPP")), 0, dR("U_TotalPP")), 2)
                cDocumento.OrdenCompra = IIf(IsDBNull(dR("U_OrdenCompra")), "", dR("U_OrdenCompra"))
                MetPago = IIf(IsDBNull(dR("U_Metodo")), "", dR("U_Metodo"))
                'MsgBox("Metodo de Pago: " + MetPago)
                Select Case MetPago
                    Case "98"
                        cDocumento.DescripcionMetodoPago = "(NO IDENTIFICADO)"
                    Case "02"
                        cDocumento.DescripcionMetodoPago = "(CHEQUE)"
                    Case "03"
                        cDocumento.DescripcionMetodoPago = "(TRANSFERENCIA ELECTRÓNICA)"
                    Case "01"
                        cDocumento.DescripcionMetodoPago = "(EFECTIVO)"
                    Case Else
                        cDocumento.DescripcionMetodoPago = ""
                End Select
                'MsgBox("Descripción Metodo de Pago: " + cDocumento.DescripcionMetodoPago)

                cDocumento.TituloDocumento = ""

                Agente = dR("SlpCode")
                DocEntry = dR("DocEntry")
                TotalDocumento = dR("DocTotal")
                cDocumento.ImporteLetra = Numero(TotalDocumento)
                PorcDescDoc = dR("DiscPrcnt") ' / 100
                UUID = IIf(IsDBNull(dR("EDocNum")), "", dR("EDocNum"))
                'MsgBox(DocEntry)
                If DocEntry = 2414 Then
                    UUID = "73D3C6EE-0251-4E17-B01D-57F4919CAEF6"
                ElseIf DocEntry = 2413 Then
                    UUID = "C57448C8-39EC-494B-9DD6-7C468D485DBC"
                End If
                FechaDocumento = dR("DocDate")


                Select Case TipoDocumento
                    Case FormatoPDF.Factura, FormatoPDF.FacturaIVA, FormatoPDF.FacturaRenta
                        cDocumento.RefPedido = IIf(IsDBNull(dR("Comments")), "", dR("Comments"))
                        cDocumento.RefPedido = cDocumento.RefPedido.Replace("Based On Sales Orders", "PEDIDO:")
                        cDocumento.RefPedido = cDocumento.RefPedido.Replace("Based On Deliveries", "ENTREGA:")
                        cDocumento.RefPedido = cDocumento.RefPedido.Replace(".", "")
                    Case Else
                        cDocumento.RefPedido = ""
                        GetFacturaAsociada(DocEntry, cDocumento.Observaciones)
                        If TipoDocumento = FormatoPDF.NotaCredito Then
                            cDocumento.TituloDocumento = "NOTA DE CRÉDITO"
                        ElseIf TipoDocumento = FormatoPDF.Devolucion Then
                            cDocumento.TituloDocumento = "DEVOLUCIÓN SOBRE VENTA"
                        End If
                End Select
            Else
                If bEsParametro Then
                    MsgBox("El número de documento no existe o no ha sido timbrado", vbExclamation)
                End If
                Return False
            End If

            If Agente > 0 Then
                sSql = "Select SlpName from OSLP where SlpCode=" + Agente.ToString
                dT = GetSQLSAP(sSql)
                If dT.Rows.Count > 0 Then
                    cDocumento.Agente = dT.Rows(0).Item("SlpName")
                End If
            Else
                cDocumento.Agente = ""
            End If

            If DocEntry > 0 Then
                Select Case TipoDocumento
                    Case FormatoPDF.Factura, FormatoPDF.FacturaIVA, FormatoPDF.FacturaRenta
                        sSql = "Select sum(isnull(Quantity,0))TotalArts from INV1 where DocEntry=" + DocEntry.ToString + " group by DocEntry"
                    Case FormatoPDF.Devolucion
                        sSql = "Select sum(isnull(Quantity,0))TotalArts from RIN1 where DocEntry=" + DocEntry.ToString + " group by DocEntry"
                    Case Else
                        sSql = ""
                End Select
                If sSql.Length > 0 Then
                    dT = GetSQLSAP(sSql)
                    If dT.Rows.Count > 0 Then
                        cDocumento.TotalArticulos = dT.Rows(0).Item("TotalArts")
                    End If
                Else
                    cDocumento.TotalArticulos = 1
                End If

                'Movimientos
                Select Case TipoDocumento
                    Case FormatoPDF.Factura, FormatoPDF.FacturaIVA, FormatoPDF.FacturaRenta
                        sSql = "Select FreeTxt, DiscPrcnt, Text from INV1 where DocEntry=" + DocEntry.ToString
                    Case FormatoPDF.Devolucion, FormatoPDF.NotaCredito
                        sSql = "Select FreeTxt, DiscPrcnt, Text from RIN1 where DocEntry=" + DocEntry.ToString
                    Case Else
                        sSql = ""
                End Select
                System.Windows.Forms.Clipboard.SetText(sSql)
                If sSql.Length > 0 Then
                    dT = GetSQLSAP(sSql)
                    For Each dR In dT.Rows
                        cMovi = New Movimiento
                        If TipoDocumento <> FormatoPDF.Devolucion Then
                            cMovi.Detalle = IIf(IsDBNull(dR("FreeTxt")), "", dR("FreeTxt"))
                            cMovi.Detalle += IIf(IsDBNull(dR("Text")), "", dR("Text"))
                        Else
                            cMovi.Detalle = "" 'GetFacturaAsociada(DocEntry) 'Tengo que sacar la factura asociada a la devolución
                        End If
                        cMovi.PorcDescto = Math.Round(PorcDescDoc, 2) 'Math.Round(PorcDescDoc * (1 + dR("DiscPrcnt") / 100) * 100, 2)
                        'Pongo esta muleta porque solo el formato de renta imprime los detalles de un movimiento
                        If cMovi.Detalle.Length > 0 Then
                            TipoDocumento = FormatoPDF.FacturaRenta
                        End If
                        cDocumento.Movimientos.Add(cMovi)
                    Next
                Else
                    cMovi = New Movimiento
                    cMovi.Detalle = "" 'GetFacturaAsociada(DocEntry) 'Tengo que sacar la factura asociada a la nota de crédito
                    cMovi.PorcDescto = 0
                    cDocumento.Movimientos.Add(cMovi)
                End If

                'Receptor
                sSql = "Select isnull(U_NoDepRef,'')U_NoDepRef from OCRD where CardCode='" + cReceptor.Codigo.ToString + "'"
                dT = GetSQLSAP(sSql)
                If dT.Rows.Count > 0 Then
                    cReceptor.RefDeposito = dT.Rows(0).Item("U_NoDepRef")
                End If
                cAddenda.Documento = cDocumento
                cAddenda.Emisor = cEmisor
                cAddenda.Receptor = cReceptor

                'Completamos la ubicación del archivo XML Origen
                Dim sXMLOrigen As String = XMLOrigen
                XMLOrigen += Year(FechaDocumento).ToString + "-" + Month(FechaDocumento).ToString("00") + "\" + cReceptor.Codigo.ToString + "\"
                Select Case TipoDocumento
                    Case FormatoPDF.Factura, FormatoPDF.FacturaIVA, FormatoPDF.FacturaRenta
                        XMLOrigen += "IN"
                    Case FormatoPDF.Devolucion, FormatoPDF.NotaCredito
                        XMLOrigen += "CM"
                    Case FormatoPDF.NotaCargo

                End Select
                XMLOrigen += "\" + UUID.ToString + ".xml"

                If Not File.Exists(XMLOrigen) Then
                    'Si no existe el XML en el mes correspondiente a la factura, lo buscamos en el siguiente mes
                    Dim FechaTMP As Date
                    If FechaDocumento.Month <> 12 Then
                        FechaTMP = DateSerial(FechaDocumento.Year, FechaDocumento.Month + 1, 1)
                    Else
                        FechaTMP = DateSerial(FechaDocumento.Year + 1, 1, 1)
                    End If
                    XMLOrigen = sXMLOrigen.ToString + Year(FechaTMP).ToString + "-" + Month(FechaTMP).ToString("00") + "\" + cReceptor.Codigo.ToString + "\"
                    Select Case TipoDocumento
                        Case FormatoPDF.Factura, FormatoPDF.FacturaIVA, FormatoPDF.FacturaRenta
                            XMLOrigen += "IN"
                        Case FormatoPDF.Devolucion, FormatoPDF.NotaCredito
                            XMLOrigen += "CM"
                        Case FormatoPDF.NotaCargo

                    End Select
                    XMLOrigen += "\" + UUID.ToString + ".xml"
                End If
                If cAddenda.Documento.Movimientos.Count > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            If bEsParametro Then
                MsgBox("Error GeneraAddendaSAP: " + ex.Message.ToString, vbExclamation)
            End If
            Return False
        End Try


    End Function

    Private Function GeneraAddendaCompac() As Boolean
        Try
            Dim cEmisor As New Emisor
            Dim cReceptor As New Receptor
            Dim cDocumento As New Documento
            Dim cMovi As Movimiento

            cDocumento.Movimientos = New List(Of Movimiento)
            cDocumento.Observaciones = New List(Of Observacion)

            Dim FolioTmp As Integer
            Dim strElementName As String

            '1) Leemos el XML y revisamos los datos que podemos obtener...
            Dim XmlRead As New XmlTextReader(XMLOrigen)
            XmlRead.WhitespaceHandling = Xml.WhitespaceHandling.Significant
            While XmlRead.Read
                If XmlRead.NodeType = Xml.XmlNodeType.Element Then
                    strElementName = XmlRead.Name
                    Select Case strElementName.ToLower
                        Case "cfdi:Comprobante".ToLower
                            FolioTmp = XmlRead.GetAttribute("folio")
                        Case "compac:DirEmp".ToLower
                            cEmisor.Telefono = XmlRead.GetAttribute("cTelefono1")
                            cEmisor.Fax = XmlRead.GetAttribute("cTelefono2")
                            cEmisor.LargaDistancia = XmlRead.GetAttribute("cTelefono3")
                            cEmisor.Web = XmlRead.GetAttribute("cDireccI01")
                        Case "compac:DirCteFis".ToLower
                            cReceptor.RefDeposito = XmlRead.GetAttribute("cDireccI01")
                        Case "compac:MGW10002".ToLower
                            cReceptor.Codigo = XmlRead.GetAttribute("cCodigoC01")
                        Case "compac:MGW10001".ToLower
                            cDocumento.Agente = XmlRead.GetAttribute("cCodigoA01")
                        Case "compac:MGW10008".ToLower
                            cDocumento.OrdenCompra = XmlRead.GetAttribute("cTextoEx03")
                            cDocumento.FechaPP = XmlRead.GetAttribute("cFechaEx01")
                            cDocumento.FechaVence = XmlRead.GetAttribute("cFechaVencimiento")
                            cDocumento.ImportePP = XmlRead.GetAttribute("cImporte01")
                            cDocumento.RefPedido = XmlRead.GetAttribute("cReferen01")
                            cDocumento.TotalArticulos = XmlRead.GetAttribute("cTotalUn01")
                        'cDocumento.Observaciones.Add(XmlRead.GetAttribute("cObserva01"))
                        Case "compac:Funciones".ToLower
                            cDocumento.ImporteLetra = XmlRead.GetAttribute("f_TOTALLETRA")
                        Case "compac:MGW10010".ToLower
                            cMovi = New Movimiento
                            cMovi.PorcDescto = XmlRead.GetAttribute("cPorcent06")
                            cMovi.Detalle = XmlRead.GetAttribute("cTextoEx01")
                            cDocumento.Movimientos.Add(cMovi)
                    End Select
                End If
            End While
            XmlRead.Close()
            cAddenda.Documento = cDocumento
            cAddenda.Emisor = cEmisor
            cAddenda.Receptor = cReceptor
            If cDocumento.Movimientos.Count > 0 Then
                If Folio = FolioTmp Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            If bEsParametro Then
                MsgBox("Error GeneraAddendaCompac: " + ex.Message.ToString, vbExclamation)
            End If
            Return False
        End Try

    End Function

    Sub GenerarAddendaXML()
        Try
            XMLAddenda = RutaXMLTemp + "Addenda" + Folio.ToString + ".xml"
            If File.Exists(XMLAddenda) Then
                File.Delete(XMLAddenda)
            End If

            Dim serialWriter As System.IO.StreamWriter
            serialWriter = New StreamWriter(XMLAddenda)
            Dim xmlWriter As New System.Xml.Serialization.XmlSerializer(GetType(Addenda))
            xmlWriter.Serialize(serialWriter, cAddenda)
            serialWriter.Close()
        Catch ex As Exception
            If bEsParametro Then
                MsgBox("Error GenerarAddendaXML: " + ex.Message.ToString, vbExclamation)
            End If
        End Try

    End Sub

    Private Sub GeneraPDF()
        Dim Etapa As String = ""

        Try
            'Nos conectamos a la BD para obtener de los parametros la direccion de servicio web que genera las facturas
            GetParametrosCrearPDF()

            Dim DestinoXML As String
            Dim FileXMLFiscal As String
            Dim FileXMLAddenda As String
            Dim FilePDF As String = ""

            If UrlGeneraPDF.Length > 0 Then
                If Directory.Exists(RutaPDFTemp) And Directory.Exists(RutaXMLTemp) Then
                    'Generamos la Addenda en la ruta XMLTemp
                    Etapa = "Generando XML"
                    GenerarAddendaXML()
                    DestinoXML = RutaXMLTemp + TipoDocumento.ToString + Folio.ToString + ".XML"
                    Etapa = "Destino XML=" + DestinoXML.ToString
                    If File.Exists(DestinoXML) Then
                        Etapa = "El XML Destino ya existía, se va a eliminar " + DestinoXML.ToString
                        File.Delete(DestinoXML)
                    End If
                    'Copiamos el XML del CFDI a \\faretdb1\inetpub\wwwroot\xml
                    Etapa = "Se va a copiar el XML a su destino." + vbCrLf + "Origen=" + XMLOrigen.ToString + vbCrLf + "Destino=" + DestinoXML.ToString
                    File.Copy(XMLOrigen, DestinoXML)

                    FileXMLFiscal = DestinoXML.Replace(RutaXMLTemp, "")
                    FileXMLAddenda = XMLAddenda.Replace(RutaXMLTemp, "")
                    FilePDF = Folio.ToString + ".PDF" 'FileXMLFiscal.ToString.Replace("XML", "PDF")

                    'Vamos a completar los parametros que se tienen que enviar al URL:
                    UrlGeneraPDF += "&xml_fiscal=" + FileXMLFiscal.ToString
                    UrlGeneraPDF += "&xml_addenda=" + FileXMLAddenda.ToString
                    UrlGeneraPDF += "&pdf=" + FilePDF.ToString
                    If bEsCopia Then
                        UrlGeneraPDF += "&copia=1"
                    End If
                    Dim tFormato As String = ""
                    Select Case TipoDocumento
                        Case FormatoPDF.Factura
                            tFormato = "f"
                        Case FormatoPDF.FacturaIVA
                            tFormato = "fiva"
                        Case FormatoPDF.FacturaRenta
                            tFormato = "frenta"
                        Case FormatoPDF.NotaCargo
                            tFormato = "ncg"
                        Case FormatoPDF.NotaCredito, FormatoPDF.Devolucion
                            tFormato = "nc"
                        Case FormatoPDF.PagoParcial
                            tFormato = "pp"
                    End Select
                    UrlGeneraPDF += "&tipo=" + tFormato.ToString

                    'Ejecutar URL: ?tipo=f/frenta/fiva/nc/d/pp&copia=0/1&xml_fiscal=test_x.xml&xml_addenda=addenda_x.xml&pdf='time'_x.pdf
                    Etapa = "URL para generar PDF:" + vbCrLf + UrlGeneraPDF
                    Dim Request As WebRequest = WebRequest.Create(UrlGeneraPDF)
                    Dim Response As WebResponse = CType(Request.GetResponse, HttpWebResponse)
                    Dim DataStream As Stream = Response.GetResponseStream
                    Dim Reader As New StreamReader(DataStream)
                    Dim ResponseFromServer As String = Reader.ReadToEnd
                    Etapa = "Respuesta del servidor :" + ResponseFromServer.ToString
                    If ResponseFromServer = 1 Then
                        'Vamos a poner un sleep si todavía no está disponible el archivo
                        For i = 1 To 30
                            If Not File.Exists(RutaPDFTemp + FilePDF) Then
                                Threading.Thread.Sleep(1000)
                            Else
                                Exit For
                            End If
                        Next

                        '''''''Etapa = "Archivo PDF por copiar: " + DestinoPDF + FilePDF
                        '''''''If File.Exists(DestinoPDF + FilePDF) Then
                        '''''''    ' MsgBox("El archivo " + DestinoPDF + FilePDF + " ya existe, se va a eliminar", vbExclamation)
                        '''''''    Etapa = "El archivo PDF Destino ya existe: " + DestinoPDF + FilePDF + " se va a eliminar"
                        '''''''    File.Delete(DestinoPDF + FilePDF)
                        '''''''    ' MsgBox("Archivo " + DestinoPDF + FilePDF + " eliminado", vbExclamation)
                        '''''''End If
                        'Copiar PDF a la carpeta DestinoPDF
                        Try
                            Etapa = "Copiando el archivo PDF." + vbCrLf + "Origen: " + RutaPDFTemp + FilePDF + vbCrLf + "Destino: " + DestinoPDF + FilePDF
                            File.Copy(RutaPDFTemp + FilePDF, DestinoPDF + FilePDF, True)
                        Catch ex As Exception
                            'MsgBox(DateDiff("n", File.GetLastWriteTime(DestinoPDF + FilePDF), Now) & vbCrLf & File.GetLastWriteTime(DestinoPDF + FilePDF) & vbCrLf & Now)
                            If DateDiff("n", File.GetLastWriteTime(DestinoPDF + FilePDF), Now) > 5 Then
                                'solo que tenga más de 5 minutos creado el archivo PDF y no se haya podido reemplezar, mostraremos error.
                                'Si no, significa se se está enviando nuevamente el correo así que el archivo creado inicialmente, nos sigue funcionando
                                MsgBox("Error GeneraPDF: " + ex.Message.ToString + vbCrLf + "Etapa: " + Etapa, vbExclamation)
                            End If
                        End Try
                    End If
                    Reader.Dispose()
                    DataStream.Dispose()
                    'Eliminamos los XML y PDF de la carpeta de temporales
                    Etapa = "Eliminando PDF Temporal: " + RutaPDFTemp + FilePDF
                    If File.Exists(RutaPDFTemp + FilePDF) Then
                        File.Delete(RutaPDFTemp + FilePDF)
                    End If
                    Etapa = "Eliminando XML Addenda Temporal: " + RutaXMLTemp + FileXMLAddenda
                    If File.Exists(RutaXMLTemp + FileXMLAddenda) Then
                        File.Delete(RutaXMLTemp + FileXMLAddenda)
                    End If
                    Etapa = "Eliminando XML Fiscal Temporal: " + RutaXMLTemp + FileXMLFiscal
                    If File.Exists(RutaXMLTemp + FileXMLFiscal) Then
                        File.Delete(RutaXMLTemp + FileXMLFiscal)
                    End If
                End If
            End If
        Catch ex As Exception
            If bEsParametro Then
                MsgBox("Error GeneraPDF: " + ex.Message.ToString + vbCrLf + Etapa, vbExclamation)
            End If
        End Try

    End Sub

    Private Sub GetFacturaAsociada(ByVal DocEntry As Integer, ByRef lstObservaciones As List(Of Observacion))
        Try
            Dim sSql As String
            Dim dT As DataTable
            sSql = "Select FolioDoc, FechaDoc from EdoCtaClientes where ObjType='13' and IdFact In(Select IdFact from EdoCtaClientes where ObjType='14' and DocEntry=" + DocEntry.ToString + ")"
            dT = GetSQLSAP(sSql)
            Dim FacturasAsociadas As New List(Of Observacion)
            Dim Obs As New Observacion
            If dT.Rows.Count > 0 Then
                Obs.Detalle = "Nota de crédito aplicada a: ".ToUpper
                FacturasAsociadas.Add(Obs)
            End If
            For Each dr In dT.Rows
                Obs = New Observacion
                Obs.Detalle = "FACTURA No. " + dr("FolioDoc").ToString + " DE FECHA " + CDate(dr("FechaDoc")).ToString("dd-MMM-yyyy").Replace(".", "").ToUpper
                FacturasAsociadas.Add(Obs)
            Next
            lstObservaciones = FacturasAsociadas
        Catch ex As Exception
            If bEsParametro Then
                MsgBox("Error GeneraPDF: " + ex.Message.ToString, vbExclamation)
            End If
        End Try

    End Sub
End Module
