Attribute VB_Name = "modConv"
Option Explicit

Public Function formatNumComprobante(pPrefijo As Integer, pNroComprob As Long)

    formatNumComprobante = Right("0000" & Trim(Str(pPrefijo)), 4) & "-" & Right("00000000" & Trim(Str(pNroComprob)), 8)
    
End Function

Public Function num2letras(ByVal pNumero As Double) As String
Dim strCentavos As String
Dim strNumero As String
Dim strUltimo As String

Dim intLargo As Integer

Dim lngNumero As Long

    strNumero = Trim(Format(pNumero, "0.00"))
    
    strCentavos = " con " & Trim(Str(Val(strNumero) * 100 - Val(Left(strNumero, Len(strNumero) - 3)) * 100)) & "/100"
    
    strNumero = Left(strNumero, Len(strNumero) - 3)
    
    intLargo = Len(strNumero)
    
    strUltimo = ""
    
    lngNumero = Val(strNumero)
    
    Select Case intLargo
        Case 1 To 3
            strUltimo = tresUltimas(lngNumero)
        Case 4 To 6
            strUltimo = tresUltimas(Int(lngNumero / 1000))
            If Right(strUltimo, 3) = "uno" Then strUltimo = Left(strUltimo, Len(strUltimo) - 1)
            strUltimo = strUltimo & " mil"
            If lngNumero - Int(lngNumero / 1000) * 1000 > 0 Then strUltimo = strUltimo & " " & Trim(tresUltimas(lngNumero - Int(lngNumero / 1000) * 1000))
        Case 7 To 9
            strUltimo = tresUltimas(Int(lngNumero / 1000000))
            If Right(strUltimo, 3) = "uno" Then strUltimo = Left(strUltimo, Len(strUltimo) - 1)
            strUltimo = strUltimo & " millon"
            If strUltimo <> " un millon" Then strUltimo = strUltimo & "es"
            lngNumero = lngNumero - Int(lngNumero / 1000000) * 1000000
            If Int(lngNumero / 1000) > 0 Then
                strUltimo = strUltimo & " " & Trim(tresUltimas(Int(lngNumero / 1000)))
                If Right(strUltimo, 3) = "uno" Then strUltimo = Left(strUltimo, Len(strUltimo) - 1)
                strUltimo = strUltimo & " mil"
            End If
            If lngNumero - Int(lngNumero / 1000) * 1000 > 0 Then strUltimo = strUltimo & " " & Trim(tresUltimas(lngNumero - Int(lngNumero / 1000) * 1000))
    End Select
    
    num2letras = Trim(strUltimo & strCentavos)
    
End Function

Private Function dosUltimas(ByVal pNumero As Integer) As String
Dim strUnidades As Variant
Dim strDecena As Variant
Dim strDecenas As Variant

Dim strUltimo As String

Dim intLargo As Integer

    intLargo = Len(Trim(Str(pNumero)))

    strUnidades = Array("uno", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve")
    strDecena = Array("diez", "once", "doce", "trece", "catorce", "quince", "dieciseis", "diecisiete", "dieciocho", "diecinueve")
    strDecenas = Array("veint", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa")
    
    strUltimo = ""

    Select Case intLargo
        Case 1
            If pNumero > 0 Then strUltimo = strUnidades(pNumero - 1)
        Case 2
            If pNumero > 9 And pNumero < 20 Then strUltimo = strDecena(pNumero - 10)
            If pNumero > 19 And pNumero < 100 Then
                strUltimo = strDecenas(Int(pNumero / 10) - 2)
                If pNumero = 20 Then strUltimo = strUltimo & "e"
                If pNumero > 20 And pNumero < 30 Then strUltimo = strUltimo & "i"
                If pNumero > 30 And pNumero < 100 And pNumero - Int(pNumero / 10) * 10 > 0 Then strUltimo = strUltimo & " y "
                If pNumero - Int(pNumero / 10) * 10 > 0 Then strUltimo = strUltimo & strUnidades(pNumero - Int(pNumero / 10) * 10 - 1)
            End If
    End Select
    
    dosUltimas = strUltimo
End Function

Private Function tresUltimas(ByVal pNumero As Integer) As String
Dim strCentenas As Variant

Dim strCentena As String
Dim strUltimo As String

    strCentenas = Array("cien", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos")
    
    strCentena = ""
        
    If pNumero > 99 Then strCentena = strCentenas(Int(pNumero / 100) - 1)
    If pNumero > 100 And pNumero < 200 Then strCentena = strCentena & "to"
    If pNumero <> Int(pNumero / 100) * 100 Then strCentena = strCentena & " "
    strUltimo = dosUltimas(pNumero - Int(pNumero / 100) * 100)
    
    tresUltimas = strCentena & strUltimo
    
End Function

Public Function num2letrasLong(ByVal pNumero As Long) As String
Dim strNumero As String
Dim strUltimo As String

Dim intLargo As Integer

Dim lngNumero As Long

    strNumero = Trim(Format(pNumero, "0"))
    
    intLargo = Len(strNumero)
    
    strUltimo = ""
    
    lngNumero = Val(strNumero)
    
    Select Case intLargo
        Case 1 To 3
            strUltimo = tresUltimas(lngNumero)
        Case 4 To 6
            strUltimo = tresUltimas(Int(lngNumero / 1000))
            If Right(strUltimo, 3) = "uno" Then strUltimo = Left(strUltimo, Len(strUltimo) - 1)
            strUltimo = strUltimo & " mil"
            If lngNumero - Int(lngNumero / 1000) * 1000 > 0 Then strUltimo = strUltimo & " " & Trim(tresUltimas(lngNumero - Int(lngNumero / 1000) * 1000))
        Case 7 To 9
            strUltimo = tresUltimas(Int(lngNumero / 1000000))
            If Right(strUltimo, 3) = "uno" Then strUltimo = Left(strUltimo, Len(strUltimo) - 1)
            strUltimo = strUltimo & " millon"
            If strUltimo <> " un millon" Then strUltimo = strUltimo & "es"
            lngNumero = lngNumero - Int(lngNumero / 1000000) * 1000000
            If Int(lngNumero / 1000) > 0 Then
                strUltimo = strUltimo & " " & Trim(tresUltimas(Int(lngNumero / 1000)))
                If Right(strUltimo, 3) = "uno" Then strUltimo = Left(strUltimo, Len(strUltimo) - 1)
                strUltimo = strUltimo & " mil"
            End If
            If lngNumero - Int(lngNumero / 1000) * 1000 > 0 Then strUltimo = strUltimo & " " & Trim(tresUltimas(lngNumero - Int(lngNumero / 1000) * 1000))
    End Select
    
    num2letrasLong = Trim(strUltimo)
    
End Function

Public Function stringAAAAMMDD2Date(pDate As String) As Date

    stringAAAAMMDD2Date = CDate(Mid(pDate, 7, 2) & "/" & Mid(pDate, 5, 2) & "/" & Mid(pDate, 1, 4))
    
End Function

Public Function stringHHMMSS2Time(pTime As String) As Date
Dim strTime As String

    strTime = Right(pTime, 6)
    stringHHMMSS2Time = CDate(Mid(strTime, 1, 2) & ":" & Mid(strTime, 3, 2) & ":" & Mid(strTime, 5, 2))
    
End Function

Public Function parseDouble(pString As String, pDecimales As Integer) As Double

    parseDouble = Val(Left(pString, Len(pString) - pDecimales)) + Val(Right(pString, pDecimales)) / Val("1" & String(pDecimales, "0"))
    
End Function

Public Function typeADO2Visual(pTypeADO As Integer) As String

    Select Case pTypeADO
        Case adChar, adLongVarChar, adLongVarWChar, adVarChar, adWChar
            typeADO2Visual = "String"
        Case adCurrency, adDecimal, adDouble, adSingle
            typeADO2Visual = "Double"
        Case adDate, adDBTimeStamp, adDBDate
            typeADO2Visual = "Date"
        Case adDBTime
            typeADO2Visual = "Time"
        Case adInteger
            typeADO2Visual = "Long"
        Case adNumeric, adSmallInt, adTinyInt
            typeADO2Visual = "Integer"
        Case adBinary
            typeADO2Visual = "Variant"
        Case Else
            typeADO2Visual = pTypeADO
    End Select
    
End Function

Public Function field2Attribute(pField As String) As String
Dim strAttribute As String

Dim blnUpper As Boolean

Dim intPosicion As Integer

    strAttribute = ""
    strAttribute = Replace(pField, "_", "")
'    blnUpper = True
'    For intPosicion = 1 To Len(pField)
'        Do While Mid(pField, intPosicion, 1) = "_"
'            blnUpper = True
'            intPosicion = intPosicion + 1
'        Loop
'        If blnUpper Then
'            strAttribute = strAttribute & UCase(Mid(pField, intPosicion, 1))
'            blnUpper = False
'        Else
'            strAttribute = strAttribute & LCase(Mid(pField, intPosicion, 1))
'        End If
'    Next intPosicion
    
    field2Attribute = strAttribute
    
End Function

Public Function parseFilename(ByVal pArchivo As String, Optional pPath As String) As String
Dim intPosicion As Integer
    
    If Trim(pArchivo) = "" Then
        parseFilename = ""
        pPath = ""
        Exit Function
    End If
    intPosicion = Len(pArchivo)
    Do While intPosicion > 0 And Mid(pArchivo, intPosicion, 1) <> "\"
        intPosicion = intPosicion - 1
    Loop
    pPath = Trim(Mid(pArchivo, 1, intPosicion))
    parseFilename = Trim(Mid(pArchivo, intPosicion + 1))

End Function
