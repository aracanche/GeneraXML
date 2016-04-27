Imports System.Data.SqlClient
Module modParametros
    Private cnSQL As SqlConnection

    Public Sub DesconectaSQL()
        Try
            cnSQL.Close()
            cnSQL = Nothing
        Catch ex As Exception

        End Try
    End Sub

    Public Function ConectaSQL() As Boolean
        cnSQL = New SqlClient.SqlConnection
        cnSQL.ConnectionString = "Data Source=FareTDB1\sap;Initial Catalog=FT" + Empresa.ToString + ";Persist Security Info=True;User ID=sa;Password=B1Admin"
        Try
            cnSQL.Open()
            Return True
        Catch ex As Exception
            MsgBox("Error al conectarse a la base de datos FT" + Empresa.ToString + " en SQL", vbExclamation)
            Return False
        End Try
    End Function

    Public Function GetSQL(ByVal sSql As String) As DataTable
        Dim dT As New DataTable
        If IsNothing(cnSQL) Then
            If Not ConectaSQL() Then
                Return dT
            End If
        End If
        Dim sqlCmd As New SqlClient.SqlCommand(sSql, cnSQL)
        Dim sqlDA As SqlClient.SqlDataAdapter
        sqlDA = New SqlClient.SqlDataAdapter(sSql, cnSQL)
        sqlDA.Fill(dT)
        Return dT
    End Function

    Public Sub GetParametrosCrearPDF()
        Dim sSQl As String
        sSQl = "Select Isnull(UrlGeneraPDF,'')UrlGeneraPDF, isnull(RutaXMLTemp,'')RutaXMLTemp, isnull(RutaPDFTemp,'')RutaPDFTemp from Parametros"
        Dim dT As DataTable
        dT = GetSQL(sSQl)
        If dT.Rows.Count > 0 Then
            UrlGeneraPDF = dT.Rows(0).Item("UrlGeneraPDF")
            RutaPDFTemp = dT.Rows(0).Item("RutaPDFTemp")
            RutaXMLTemp = dT.Rows(0).Item("RutaXMLTemp")
            If Not RutaPDFTemp.EndsWith("\") Then
                RutaPDFTemp += "\"
            End If
            If Not RutaXMLTemp.EndsWith("\") Then
                RutaXMLTemp += "\"
            End If
        End If
    End Sub
End Module
