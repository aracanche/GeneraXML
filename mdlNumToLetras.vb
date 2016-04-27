Module mdlNumToLetras
    Public Function Numero(ByVal Num As Double)
        '****************************************************************************
        '* Objetivo: Escribe un Cantidad con Letra
        '* Fecha: 20/01/99
        '****************************************************************************

        If Num <= 9999999.99 Then
            'Declaración de Variables
            Dim v1 As String, v2 As String, v3 As String, v4 As String
            Dim Largo As Integer, Cadena As String, numcad As String

            'Completa la Cantidad a 13 espacios
            Largo = 12 - Len(Format(Num, "#,###,###.00"))
            Cadena = Space(Largo) + Format(Num, "#,###,###.00")

            'Extrae los Valores a Utilizar
            v1 = Mid(Cadena, 1, 1)   'Millares
            v2 = Mid(Cadena, 3, 3)   'Miles
            v3 = Mid(Cadena, 7, 3)   'Cientos
            v4 = Mid(Cadena, 11, 2)  'Centavos
            'Llama a funciones de formatos
            'Set rsCur = db.Execute("SELECT * FROM claves WHERE clave = '" & lblMoneda.Caption & "' AND tipo = 'M'")
            numcad = "(" + unidadx(v1, 1) + Ciento(v2, 1) + Decenas(v2, 1) + unidadx(v2, 2) + Ciento(v3, 2) + Decenas(v3, 2) + unidadx(v3, 3)
            'numcad = numcad + UCase(rsCur!descripcion) + " " + v4 + "/100" + IIf(lblMoneda.Caption = "MXP", " M.N.", "") + ")"
            'If lblMoneda.Caption <> "" And rsCur.RecordCount Then
            '    numcad = numcad + UCase(rsCur!Descripcion) + " "
            'End If
            numcad = numcad + "PESOS " + v4 + "/100" + " M.N." + ")"
            'Asigna la cadena a la Función Numero
            Numero = numcad
        Else
            Numero = ""
        End If
    End Function

    Function Decenas(ByVal can, ByVal Tipo)
        '****************************************************************************
        '* Objetivo: Obtiene Valor Decenas
        '* Fecha: 20/01/99
        '****************************************************************************
        'Declaración de Variables
        Dim v As Single, X As Single, Y As Single, d As String
        d = ""
        v = Val(Mid(can, 2, 1)) 'Digito de Decenas
        Y = Val(Right(can, 2))  'Valor Decenas
        X = Val(Right(can, 1))  'Valor unidadxes

        Select Case v
            Case 2
                d = IIf(Val(Y) = 20, "VEINTE ", "VEINTI")
            Case 3
                d = "TREINTA "
            Case 4
                d = "CUARENTA "
            Case 5
                d = "CINCUENTA "
            Case 6
                d = "SESENTA "
            Case 7
                d = "SETENTA "
            Case 8
                d = "OCHENTA "
            Case 9
                d = "NOVENTA "
        End Select

        If X > 0 And v > 2 Then
            d = d + "Y "
        End If
        Decenas = d
    End Function

    Function Ciento(ByVal can, ByVal Tipo)
        '****************************************************************************
        '* Objetivo: Obtiene Valor Centenas
        '* Fecha: 20/01/99
        '****************************************************************************
        'Declaración de Variables
        Dim d As String
        d = ""
        Select Case Val(can)
            Case 100
                d = "CIEN "
            Case 101 To 199
                d = "CIENTO "
            Case 200 To 299
                d = "DOSCIENTOS "
            Case 300 To 399
                d = "TRESCIENTOS "
            Case 400 To 499
                d = "CUATROCIENTOS "
            Case 500 To 599
                d = "QUINIENTOS "
            Case 600 To 699
                d = "SEISCIENTOS "
            Case 700 To 799
                d = "SETECIENTOS "
            Case 800 To 899
                d = "OCHOCIENTOS "
            Case 900 To 999
                d = "NOVECIENTOS "
        End Select
        Ciento = d
    End Function

    Function unidadx(ByVal can, ByVal Tipo)
        '****************************************************************************
        '* Objetivo: Selecciona las unidadxes
        '* Fecha: 20/01/99
        '****************************************************************************
        'Declaración de Variables
        Dim v As String, X As Single, Y As Single, z As Single
        Dim d As String
        d = ""
        v = Right(can, 1)
        X = Val(Right(can, 2))
        Y = Val(Mid(can, 2, 1))

        Select Case v
            Case "0" And X > 0
                If Tipo = 2 Then d = "MIL " Else If Tipo = 1 Then d = "CERO "
            Case "1"
                Select Case Tipo
                    Case 1
                        d = "UN MILLON "
                    Case 2 And Y <> 1
                        d = "UN MIL "
                    Case 3 And Y <> 1
                        d = "UN "
                End Select
            Case "2"
                Select Case Tipo
                    Case 1
                        d = "DOS MILLONES "
                    Case 2 And Y <> 1
                        d = "DOS MIL "
                    Case 3 And Y <> 1
                        d = "DOS "
                End Select
            Case "3"
                Select Case Tipo
                    Case 1
                        d = "TRES MILLONES "
                    Case 2 And Y <> 1
                        d = "TRES MIL "
                    Case 3 And Y <> 1
                        d = "TRES "
                End Select
            Case "4"
                Select Case Tipo
                    Case 1
                        d = "CUATRO MILLONES "
                    Case 2 And Y <> 1
                        d = "CUATRO MIL "
                    Case 3 And Y <> 1
                        d = "CUATRO "
                End Select
            Case "5"
                Select Case Tipo
                    Case 1
                        d = "CINCO MILLONES "
                    Case 2 And Y <> 1
                        d = "CINCO MIL "
                    Case 3 And Y <> 1
                        d = "CINCO "
                End Select
            Case "6"
                Select Case Tipo
                    Case 1
                        d = "SEIS MILLONES "
                    Case 2 And Y <> 1
                        d = "SEIS MIL "
                    Case 3 And Y <> 1
                        d = "SEIS "
                End Select
            Case "7"
                Select Case Tipo
                    Case Is = 1
                        d = "SIETE MILLONES "
                    Case 2 And Y <> 1
                        d = "SIETE MIL "
                    Case 3 And Y <> 1
                        d = "SIETE "
                End Select
            Case "8"
                Select Case Tipo
                    Case 1
                        d = "OCHO MILLONES "
                    Case 2 And Y <> 1
                        d = "OCHO MIL "
                    Case 3 And Y <> 1
                        d = "OCHO "
                End Select
            Case "9"
                Select Case Tipo
                    Case 1
                        d = "NUEVE MILLONES "
                    Case 2 And Y <> 1
                        d = "NUEVE MIL "
                    Case 3 And Y <> 1
                        d = "NUEVE "
                End Select
        End Select
        z = 0
        Select Case X
            Case 10
                d = "DIEZ "
                z = 1
            Case 11
                d = "ONCE "
                z = 1
            Case 12
                d = "DOCE "
                z = 1
            Case 13
                d = "TRECE "
                z = 1
            Case 14
                d = "CATORCE "
                z = 1
            Case 15
                d = "QUINCE "
                z = 1
            Case 16
                d = "DIECISEIS "
                z = 1
            Case 17
                d = "DIECISIETE "
                z = 1
            Case 18
                d = "DIECIOCHO "
                z = 1
            Case 19
                d = "DIECINUEVE "
                z = 1
        End Select
        If z = 1 Or X = 0 Then
            If Tipo = 2 And Val(can) > 0 Then
                d = d + "MIL "
            End If
        End If
        unidadx = d
    End Function
End Module
