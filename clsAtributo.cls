VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAtributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vAtributo As String
Private vTipo As String
Private vPrimaria As Boolean
Private vFind As Boolean
Private vNullable As Boolean
Private vAutoIncrement As Boolean

Private Sub Class_Initialize()

    vAtributo = ""
    vTipo = ""
    vPrimaria = False
    vFind = False
    vNullable = False
    vAutoIncrement = False
    
End Sub

Public Property Get atributo() As Variant

    atributo = vAtributo
    
End Property

Public Property Let atributo(ByVal vNewValue As Variant)

    vAtributo = vNewValue
    
End Property

Public Property Get tipo() As Variant

    tipo = vTipo
    
End Property

Public Property Let tipo(ByVal vNewValue As Variant)

    vTipo = vNewValue
    
End Property

Public Property Get primaria() As Variant

    primaria = vPrimaria
    
End Property

Public Property Let primaria(ByVal vNewValue As Variant)

    vPrimaria = vNewValue
    
End Property

Public Property Get find() As Variant

    find = vFind
    
End Property

Public Property Let find(ByVal vNewValue As Variant)

    vFind = vNewValue
    
End Property

Public Property Get nullable() As Variant

    nullable = vNullable
    
End Property

Public Property Let nullable(ByVal vNewValue As Variant)

    vNullable = vNewValue
    
End Property

Public Property Get autoIncrement() As Variant

    autoIncrement = vAutoIncrement
    
End Property

Public Property Let autoIncrement(ByVal vNewValue As Variant)

    vAutoIncrement = vNewValue
    
End Property

Public Property Get typeCreation() As Variant

    typeCreation = vTipo
    If vNullable Then typeCreation = "Variant"
    
End Property

Public Function existPrimary(pAtributos As Collection) As Boolean

    existPrimary = Not (makePrimaryCollection(pAtributos).Count = 0)
    
End Function

Public Function isPrimaryAutoIncrement(pAtributos As Collection) As Boolean
Dim objAtributo As clsAtributo

    isPrimaryAutoIncrement = False
    
    For Each objAtributo In makePrimaryCollection(pAtributos)
        If objAtributo.primaria And objAtributo.autoIncrement Then
            isPrimaryAutoIncrement = True
            Exit Function
        End If
    Next
    
End Function

Public Function makePrimaryCollection(pAtributos As Collection, Optional pLess As Variant) As Collection
Dim objAtributo As clsAtributo

Dim colTemp As New Collection
Dim colPrimary As New Collection

Dim intCiclo As Integer

    For Each objAtributo In pAtributos
        If objAtributo.primaria Then colTemp.Add objAtributo
    Next
    
    If IsMissing(pLess) Then
        Set makePrimaryCollection = colTemp
        
        Exit Function
    End If
    
    For intCiclo = 1 To colTemp.Count - pLess
        colPrimary.Add colTemp(intCiclo)
    Next intCiclo
    
    Set makePrimaryCollection = colPrimary
    
End Function

Public Function makeNoPrimaryCollection(pAtributos As Collection, Optional update As Boolean = False) As Collection
Dim objAtributo As clsAtributo

Dim colPrimary As New Collection

    For Each objAtributo In pAtributos
        If Not objAtributo.primaria Then colPrimary.Add objAtributo
    Next
    
    Set makeNoPrimaryCollection = colPrimary
    
End Function

Public Function makeFindCollection(pAtributos As Collection) As Collection
Dim objAtributo As clsAtributo

Dim colFind As New Collection

    For Each objAtributo In pAtributos
        If objAtributo.find Then colFind.Add objAtributo
    Next
    
    Set makeFindCollection = colFind
    
End Function

Public Function lastPrimary(pAtributos As Collection) As clsAtributo
Dim colPrimary As New Collection

    Set colPrimary = makePrimaryCollection(pAtributos)
    
    Set lastPrimary = colPrimary(colPrimary.Count)
    
End Function

Public Function firstNoPrimary(pAtributos As Collection) As clsAtributo
Dim colNoPrimary As New Collection

    Set firstNoPrimary = Nothing

    Set colNoPrimary = makeNoPrimaryCollection(pAtributos)
    
    If colNoPrimary.Count > 0 Then Set firstNoPrimary = colNoPrimary(1)
    
End Function

Public Function makeWherePrimary(pAtributos As Collection, pIndentPosition As Integer, Optional pLess As Variant) As Collection
Dim strLinea As String

Dim blnPrimera As Boolean

Dim objAtributo As clsAtributo

Dim colWhere As New Collection
Dim colPrimary As New Collection

    blnPrimera = True
    
    Set colPrimary = makePrimaryCollection(pAtributos, pLess)
    
    For Each objAtributo In colPrimary
        If blnPrimera Then
            strLinea = Space(pIndentPosition) & "sql = sql & " & Chr(34) & " WHERE "
            blnPrimera = False
        Else
            strLinea = Space(pIndentPosition) & "sql = sql & " & Chr(34) & " AND "
        End If
        strLinea = strLinea & Chr(34) & " & ddl." & firstLower(objAtributo.atributo) & " & " & Chr(34) & " = " & objAtributo.makeSQL
        
        colWhere.Add strLinea
    Next
    
    Set makeWherePrimary = colWhere
    
End Function

Public Function makeSQL(Optional update As Boolean = False) As String

    If vTipo <> "Date" And vTipo <> "Time" And vNullable Then
        makeSQL = Chr(34) & " & validateNullField(v" & firstUpper(vAtributo) & ")"
        Exit Function
    End If
    
    Select Case vTipo
        Case "Integer", "Long"
            makeSQL = Chr(34) & " & v" & firstUpper(vAtributo)
        Case "Date"
            makeSQL = Chr(34) & " & fechaDB(v" & firstUpper(vAtributo) & ")"
        Case "Time"
            makeSQL = Chr(34) & " & horaDB(v" & firstUpper(vAtributo) & ")"
        Case "String"
            makeSQL = "'" & Chr(34) & " & ReplaceSpecialSymbols(v" & firstUpper(vAtributo) & ") & " & Chr(34) & "'" & Chr(34)
        Case "Double", "Currency"
            makeSQL = "'" & Chr(34) & " & Replace(v" & firstUpper(vAtributo) & ", " & Chr(34) & "," & Chr(34) & ", " & Chr(34) & "." & Chr(34) & ") & " & Chr(34) & "'" & Chr(34)
    End Select
    
End Function

Public Function makeSetNoPrimary(pAtributos As Collection, pIndentPosition As Integer) As Collection
Dim strLinea As String

Dim blnPrimera As Boolean

Dim objAtributo As clsAtributo

Dim colWhere As New Collection
Dim colNoPrimary As New Collection

    blnPrimera = True
    
    Set colNoPrimary = makeNoPrimaryCollection(pAtributos)
    
    For Each objAtributo In colNoPrimary
        If Not objAtributo.autoIncrement Then
            If blnPrimera Then
                strLinea = Space(pIndentPosition) & "sql = sql & " & Chr(34) & " SET "
                blnPrimera = False
            Else
                strLinea = Space(pIndentPosition) & "sql = sql & " & Chr(34) & ", "
            End If
            strLinea = strLinea & objAtributo.makeFieldEqualValue
            
            colWhere.Add strLinea
        End If
    Next
    
    Set makeSetNoPrimary = colWhere
    
End Function

Public Function makeSetNoPrimarySmart(pAtributos As Collection, pIndentPosition As Integer, Optional update As Boolean = False) As Collection
Dim strLinea As String

Dim objAtributo As clsAtributo

Dim colWhere As New Collection
Dim colNoPrimary As New Collection

    Set colNoPrimary = makeNoPrimaryCollection(pAtributos, update)
    
    For Each objAtributo In colNoPrimary
        If Not objAtributo.autoIncrement Then
            strLinea = Space(pIndentPosition) & "If v" & firstUpper(objAtributo.atributo) & " <> clon." & objAtributo.atributo & " Or IsNull(v" & firstUpper(objAtributo.atributo) & ") <> IsNull(clon." & objAtributo.atributo & ") Then" & vbCrLf
            strLinea = strLinea & Space(pIndentPosition) & "    If strSET <> " & Chr(34) & Chr(34) & " Then strSET = strSET & " & Chr(34) & ", " & Chr(34) & vbCrLf
            strLinea = strLinea & Space(pIndentPosition) & "    strSET = strSET & " & Chr(34) & " " & objAtributo.makeFieldEqualValue & vbCrLf
            strLinea = strLinea & Space(pIndentPosition) & "    blnChange = True" & vbCrLf
            strLinea = strLinea & Space(pIndentPosition) & "End If"
            
            colWhere.Add strLinea
        End If
    Next
    
    Set makeSetNoPrimarySmart = colWhere
    
End Function

Public Function makeSetPrimary(pAtributos As Collection, pIndentPosition As Integer) As Collection
Dim strLinea As String

Dim objAtributo As clsAtributo

Dim colWhere As New Collection
Dim colPrimary As New Collection

    Set colPrimary = makePrimaryCollection(pAtributos)
    
    For Each objAtributo In colPrimary
        If Not objAtributo.autoIncrement Then
            strLinea = Space(pIndentPosition) & "sql = sql & " & Chr(34) & ", "
            strLinea = strLinea & objAtributo.makeFieldEqualValue
            
            colWhere.Add strLinea
        End If
    Next
    
    Set makeSetPrimary = colWhere
    
End Function

Public Function makeFieldEqualValue(Optional update As Boolean = False) As String

    makeFieldEqualValue = Chr(34) & " & ddl." & firstLower(Me.atributo) & " & " & Chr(34) & " = " & Me.makeSQL(update)
    
End Function
