Attribute VB_Name = "CompruebaCCC"
'-- Esta librería contiene un conjunto de funciones de utilidad general
Public Function Comprueba_CC(CC As String) As Boolean
    Dim ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim I, i2, i3, i4 As Integer
    Dim NumCC As String ' Número de cuenta propiamente dicho
    '-- Esta función comprueba la corrección de un número de cuenta pasado en CC
    If Len(CC) <> 20 Then Exit Function '-- Las cuentas deben contener 20 dígitos en total
    
    
    '-- Calculamos el primer dígito de control
    I = Val(Mid(CC, 1, 1)) * 4
    I = I + Val(Mid(CC, 2, 1)) * 8
    I = I + Val(Mid(CC, 3, 1)) * 5
    I = I + Val(Mid(CC, 4, 1)) * 10
    I = I + Val(Mid(CC, 5, 1)) * 9
    I = I + Val(Mid(CC, 6, 1)) * 7
    I = I + Val(Mid(CC, 7, 1)) * 3
    I = I + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 9, 1)) Then Exit Function '-- El primer dígito de control no coincide
    '-- Calculamos el segundo dígito de control
    I = Val(Mid(CC, 11, 1)) * 1
    I = I + Val(Mid(CC, 12, 1)) * 2
    I = I + Val(Mid(CC, 13, 1)) * 4
    I = I + Val(Mid(CC, 14, 1)) * 8
    I = I + Val(Mid(CC, 15, 1)) * 5
    I = I + Val(Mid(CC, 16, 1)) * 10
    I = I + Val(Mid(CC, 17, 1)) * 9
    I = I + Val(Mid(CC, 18, 1)) * 7
    I = I + Val(Mid(CC, 19, 1)) * 3
    I = I + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 10, 1)) Then Exit Function '-- El segundo dígito de control no coincide
    '-- Si llega aquí ambos figitos de control son correctos
    Comprueba_CC = True
End Function


'---- Añade Laura: 04/10/05
Public Function Comprueba_CuentaBan(CC As String) As Boolean
    'Validar que la cuenta bancaria es correcta
    If Trim(CC) <> "" Then
        If Not Comprueba_CC(CC) Then
            MsgBox "La cuenta bancaria no es correcta", vbInformation
        End If
    End If
End Function
'------------------------------


'[Monica]20/11/2013:
Public Function Comprueba_CC_IBAN(CC As String, IBAN As String) As Boolean
    Dim ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim I, i2, i3, i4 As Integer
    Dim NumCC As String ' Número de cuenta propiamente dicho
    '-- Esta función comprueba la corrección de un número de cuenta pasado en CC
    
    
    If Len(IBAN) <> 4 Then Exit Function '-- Las cuentas deben contener 20 dígitos en total
    
    
    '-- Calculamos el primer dígito de control
    I = Val(Mid(CC, 1, 1)) * 4
    I = I + Val(Mid(CC, 2, 1)) * 8
    I = I + Val(Mid(CC, 3, 1)) * 5
    I = I + Val(Mid(CC, 4, 1)) * 10
    I = I + Val(Mid(CC, 5, 1)) * 9
    I = I + Val(Mid(CC, 6, 1)) * 7
    I = I + Val(Mid(CC, 7, 1)) * 3
    I = I + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 9, 1)) Then Exit Function '-- El primer dígito de control no coincide
    '-- Calculamos el segundo dígito de control
    I = Val(Mid(CC, 11, 1)) * 1
    I = I + Val(Mid(CC, 12, 1)) * 2
    I = I + Val(Mid(CC, 13, 1)) * 4
    I = I + Val(Mid(CC, 14, 1)) * 8
    I = I + Val(Mid(CC, 15, 1)) * 5
    I = I + Val(Mid(CC, 16, 1)) * 10
    I = I + Val(Mid(CC, 17, 1)) * 9
    I = I + Val(Mid(CC, 18, 1)) * 7
    I = I + Val(Mid(CC, 19, 1)) * 3
    I = I + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    If i4 <> Val(Mid(CC, 10, 1)) Then Exit Function '-- El segundo dígito de control no coincide
    '-- Si llega aquí ambos figitos de control son correctos
    Comprueba_CC_IBAN = True
End Function


Public Function Calculo_CC_IBAN(CC As String, IBAN As String) As String
    Dim ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim I, i2, i3, i4 As Integer
    Dim NumCC As String ' Número de cuenta propiamente dicho
    '-- Esta función comprueba la corrección de un número de cuenta pasado en CC
    Dim vIban As String
    
    Dim v1 As String
    Dim v2 As String
    Dim n1 As Integer
    Dim n2 As String
    
    Resul = 0
    
    If Len(CC) <> 20 Then Exit Function
    If Len(IBAN) = 0 Then
        vIban = "ES"
    Else
        vIban = IBAN
    End If
    
    
    If IsNumeric(Mid(vIban, 1, 2)) Then
        Exit Function
    Else
        v1 = Mid(UCase(vIban), 1, 1)
        v2 = Mid(UCase(vIban), 2, 1)
        If Asc(v1) >= 65 And Asc(v1) <= 90 And Asc(v2) >= 65 And Asc(v2) <= 90 Then
            n1 = ValorLetra(v1)
            n2 = ValorLetra(v2)
        End If
        CC = CC & n1 & n2
        'resul = 98 - (CDbl(CC) \ 97)
    End If
    
    'Calculo_CC_IBAN = Mid(vIban, 1, 2) & Format(resul, "00")
    
    ' Calculo a tramos pq no se puede una longitud tan larga
    cc1 = Mid(CC, 1, 9)
    cc2 = Mid(CC, 10, Len(CC) - 9)
    For I = 1 To 4
        dig1 = cc1 Mod 97
        cc1 = dig1 & cc2
        If cc2 = "" Then Exit For
        If Len(cc1) > 9 Then
            cc2 = Mid(cc1, 10, Len(cc1))
            cc1 = Mid(cc1, 1, 9)
        Else
            cc2 = ""
            'cc1 como está
        End If
    Next I
    Resul = 98 - dig1
    Calculo_CC_IBAN = Mid(vIban, 1, 2) & Format(Resul, "00")
    
End Function

Private Function ValorLetra(LEtra As String) As Byte
Dim Valor As Byte

    If Asc(LEtra) >= 65 And Asc(LEtra) <= 90 Then
        Valor = Asc(LEtra) - 55
    End If
    
    ValorLetra = Valor
    
End Function



Public Function DigitoControlCorrecto(CC As String) As String
    Dim ent As String ' Entidad
    Dim Suc As String ' Oficina
    Dim DC As String ' Digitos de control
    Dim I, i2, i3, i4 As Integer
    Dim NumCC As String ' Número de cuenta propiamente dicho
    '-- Esta función comprueba la corrección de un número de cuenta pasado en CC
    If Len(CC) <> 20 Then Exit Function '-- Las cuentas deben contener 20 dígitos en total
    
    
    '-- Calculamos el primer dígito de control
    I = Val(Mid(CC, 1, 1)) * 4
    I = I + Val(Mid(CC, 2, 1)) * 8
    I = I + Val(Mid(CC, 3, 1)) * 5
    I = I + Val(Mid(CC, 4, 1)) * 10
    I = I + Val(Mid(CC, 5, 1)) * 9
    I = I + Val(Mid(CC, 6, 1)) * 7
    I = I + Val(Mid(CC, 7, 1)) * 3
    I = I + Val(Mid(CC, 8, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    DC = i4
    '-- Calculamos el segundo dígito de control
    I = Val(Mid(CC, 11, 1)) * 1
    I = I + Val(Mid(CC, 12, 1)) * 2
    I = I + Val(Mid(CC, 13, 1)) * 4
    I = I + Val(Mid(CC, 14, 1)) * 8
    I = I + Val(Mid(CC, 15, 1)) * 5
    I = I + Val(Mid(CC, 16, 1)) * 10
    I = I + Val(Mid(CC, 17, 1)) * 9
    I = I + Val(Mid(CC, 18, 1)) * 7
    I = I + Val(Mid(CC, 19, 1)) * 3
    I = I + Val(Mid(CC, 20, 1)) * 6
    i2 = Int(I / 11)
    i3 = I - (i2 * 11)
    i4 = 11 - i3
    Select Case i4
        Case 11
            i4 = 0
        Case 10
            i4 = 1
    End Select
    DC = DC & i4
    '-- Si llega aquí ambos figitos de control son correctos
    DigitoControlCorrecto = DC
End Function



