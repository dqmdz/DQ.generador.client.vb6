VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private vAtributos As Collection
Private vPath As String
Private vFile As String
Private vClassName As String
Private vEndpoint As String
Private vUrl As String

Private Sub Class_Initialize()

    Set vAtributos = New Collection
    
End Sub

Public Sub clean()

    Call Class_Initialize
    
End Sub

Public Property Get atributos() As Collection

    Set atributos = vAtributos
    
End Property

Public Property Get path() As Variant

    path = vPath
    
End Property

Public Property Let path(ByVal vNewValue As Variant)

    vPath = vNewValue
    
End Property

Public Property Get file() As Variant

    file = vFile
    
End Property

Public Property Let file(ByVal vNewValue As Variant)

    vFile = vNewValue
    
End Property

Public Property Get className() As Variant

    className = vClassName
    
End Property

Public Property Let className(ByVal vNewValue As Variant)

    vClassName = vNewValue
    
End Property

Public Property Get endpoint() As Variant

    endpoint = vEndpoint
    
End Property

Public Property Let endpoint(ByVal vNewValue As Variant)

    vEndpoint = vNewValue
    
End Property

Public Property Get url() As Variant

    url = vUrl
    
End Property

Public Property Let url(ByVal vNewValue As Variant)

    vUrl = vNewValue
    
End Property

Public Sub writeHeader()
Dim objClass As New clsClass

    Open vPath & vFile For Output As #1
    
    objClass.writeHeader 1, vClassName
    
End Sub

Public Sub closeFile()

    Close #1
    
End Sub

Public Sub writeAttributes()
Dim objAtributo As clsAtributo

    For Each objAtributo In vAtributos
        Print #1, "Private v" & objAtributo.atributo & " As " & objAtributo.typeCreation
    Next
    Print #1, ""

End Sub

Public Sub writeInitialize()
Dim objAtributo As New clsAtributo

Dim objTipo As New clsTipo

    printHeader "Inicializacion de la Clase"
    Print #1, "Private Sub class_Initialize()"
    Print #1, ""
    
    If objAtributo.existPrimary(vAtributos) Then
        For Each objAtributo In objAtributo.makePrimaryCollection(vAtributos)
            Print #1, "    v" & objAtributo.atributo & " = " & objTipo.initialValue(objAtributo)
        Next
    End If
    
    For Each objAtributo In objAtributo.makeNoPrimaryCollection(vAtributos)
        Print #1, "    v" & objAtributo.atributo & " = " & objTipo.initialValue(objAtributo)
    Next
    
    Print #1, ""
    Print #1, "End Sub"
    Print #1, ""
    
    printHeader "Limpia todos los atributos"
    Print #1, "Public Sub clean()"
    Print #1, ""
    Print #1, "    Call class_Initialize"
    Print #1, ""
    Print #1, "End Sub"
    Print #1, ""
    
End Sub

Public Sub writeGetterSetter()
Dim objAtributo As New clsAtributo

    ' Genera atributos
    For Each objAtributo In vAtributos
        printHeader "Getter " & firstLower(objAtributo.atributo)
        Print #1, "Public Property Get " & firstLower(objAtributo.atributo) & "() As Variant"
        Print #1, ""
        Print #1, "    " & firstLower(objAtributo.atributo) & " = v" & firstUpper(objAtributo.atributo)
        Print #1, ""
        Print #1, "End Property"
        Print #1, ""
        printHeader "Setter " & firstLower(objAtributo.atributo)
        Print #1, "Public Property Let " & firstLower(objAtributo.atributo) & "(ByVal vNewValue As Variant)"
        Print #1, ""
        Print #1, "On Error Resume Next"
        Print #1, ""
        Print #1, "    v" & firstUpper(objAtributo.atributo) & " = vNewValue"
        Print #1, ""
        Print #1, "End Property"
        Print #1, ""
    Next
    
    printHeader "Atributo key para usar como clave de colecciones"
    Print #1, "Public Property Get key() As String"
    Print #1, ""
    Print #1, "    key = " & Chr(34) & "k." & Chr(34) & " &"
    Print #1, ""
    Print #1, "End Property"
    Print #1, ""

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Getter comboText"
        Print #1, "Public Property Get comboText() As Variant"
        Print #1, ""
        Print #1, "    comboText = v" & firstUpper(objAtributo.firstNoPrimary(vAtributos).atributo)
        Print #1, ""
        Print #1, "End Property"
        Print #1, ""
        printHeader "Getter comboKey"
        Print #1, "Public Property Get comboKey() As Variant"
        Print #1, ""
        Print #1, "    comboKey = v" & firstUpper(objAtributo.lastPrimary(vAtributos).atributo)
        Print #1, ""
        Print #1, "End Property"
        Print #1, ""
        printHeader "Getter listText"
        Print #1, "Public Property Get listText() As Variant"
        Print #1, ""
        Print #1, "    listText = v" & firstUpper(objAtributo.firstNoPrimary(vAtributos).atributo)
        Print #1, ""
        Print #1, "End Property"
        Print #1, ""
        printHeader "Getter listKey"
        Print #1, "Public Property Get listKey() As Variant"
        Print #1, ""
        Print #1, "    listKey = v" & firstUpper(objAtributo.lastPrimary(vAtributos).atributo)
        Print #1, ""
        Print #1, "End Property"
        Print #1, ""
    End If

End Sub

Public Sub writeFillObject()
Dim objAtributo As clsAtributo

    If Trim(vEndpoint) <> "" Then
        ' Genera fillObject
        printHeader "Carga los campos del registro en los atributos de la clase"
        Print #1, "Public Sub fillObject(json_object As Dictionary)"
        Print #1, ""
        Print #1, "On Error Resume Next"
        Print #1, ""
        
        For Each objAtributo In vAtributos
            If objAtributo.tipo = "Date" Then
                Print #1, "    " & firstLower(objAtributo.atributo) & " = modConv.JSON2Date(json_object.item(" & Chr(34) & firstLower(objAtributo.atributo) & Chr(34) & "))"
            Else
                Print #1, "    " & firstLower(objAtributo.atributo) & " = json_object.item(" & Chr(34) & firstLower(objAtributo.atributo) & Chr(34) & ")"
            End If
        Next
        
        Print #1, ""
        Print #1, "End Sub"
        Print #1, ""
    End If

End Sub

Public Sub writeClone()
Dim objAtributo As clsAtributo

    printHeader "Clona el objeto this y devuelve uno nuevo"
    Print #1, "Public Function clone() As " & vClassName
    Print #1, "Dim object As New " & vClassName
    Print #1, ""
    
    For Each objAtributo In vAtributos
        Print #1, "    object." & firstLower(objAtributo.atributo) & " = v" & objAtributo.atributo
    Next
    
    Print #1, ""
    Print #1, "    Set clone = object"
    Print #1, ""
    Print #1, "End Function"
    Print #1, ""

End Sub

Public Sub writeFindREST()
Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Consulta la DB con un SQL buscando 1 registro"
        Print #1, "Private Function findREST(json_object As Dictionary) As " & vClassName
        Print #1, "Dim " & LCase(vEndpoint) & " As New " & vClassName
        Print #1, ""
        Print #1, "    If json_object.item(" & Chr(34) & "status" & Chr(34) & ") <> " & Chr(34) & Chr(34) & " Then"
        Print #1, "        Set findREST = new " & vClassName
        Print #1, "        Exit Function"
        Print #1, "    End If"
        Print #1, ""
        Print #1, "    " & vEndpoint & ".fillObject json_object"
        Print #1, ""
        Print #1, "    Set findREST = " & LCase(vEndpoint)
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If
    
End Sub

Public Sub writeFindByPrimaryKey()
Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Realiza un find sobre los atributos de la Primary Key"
        Print #1, "Public Function findBy" & vEndpoint & "Id(" & vEndpoint & "Id As Long) As " & vClassName
        Print #1, "Dim request As MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "Dim url As String"
        Print #1, ""
        Print #1, "    Set request = New MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "    url = " & vUrl & " & " & Chr(34) & "/" & LCase(vEndpoint) & "/" & Chr(34) & " & " & LCase(vEndpoint) & "Id"
        Print #1, ""
        Print #1, "    request.Open " & Chr(34) & "GET" & Chr(34) & ", url"
        Print #1, "    request.send"
        Print #1, ""
        Print #1, "    Set findBy" & vEndpoint & "Id = findREST(modJsonConverter.ParseJson(request.responseText))"
        Print #1, ""
        Print #1, "    Set request = Nothing"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If

End Sub

Public Sub writeDelete()
Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Elimina 1 registro de la DB según la Primary Key"
        Print #1, "Public Function deleteBy" & vEndpoint & "Id(" & LCase(vEndpoint) & "Id As Long) As Boolean"
        Print #1, "Dim request As MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "Dim url As String"
        Print #1, ""
        Print #1, "    deleteBy" & vEndpoint & "Id = False"
        Print #1, ""
        Print #1, "    Set request = New MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "    url = " & vUrl & " & " & Chr(34) & "/" & LCase(vEndpoint) & "/" & Chr(34) & " & " & LCase(vEndpoint) & "Id"
        Print #1, ""
        Print #1, "    request.Open " & Chr(34) & "DELETE" & Chr(34) & ", url"
        Print #1, "    request.send"
        Print #1, ""
        Print #1, "    If request.responseText = " & Chr(34) & Chr(34) & " Then deleteBy" & vEndpoint & "Id = True"
        Print #1, ""
        Print #1, "    Set request = Nothing"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If

End Sub

Public Sub writeSave()
Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Graba un registro verificando si debe agregarlo o actualizarlo"
        Print #1, "Public Function save(" & LCase(vEndpoint) & " As " & vClassName & ") As " & vClassName
        Print #1, ""
        Print #1, "    If exist(" & LCase(vEndpoint) & "." & LCase(vEndpoint) & "Id) Then"
        Print #1, "        Set save = update(" & LCase(vEndpoint) & ", " & LCase(vEndpoint) & "." & LCase(vEndpoint) & "Id)"
        Print #1, "    Else"
        Print #1, "        Set save = add(" & LCase(vEndpoint) & ")"
        Print #1, "    End If"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If
    
End Sub

Public Sub writeExist()
Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Verifica si el registro asociado a la Primary Key existe"
        Print #1, "Public Function exist(" & LCase(vEndpoint) & "Id As Long) As Boolean"
        Print #1, "Dim request As MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "Dim url As String"
        Print #1, ""
        Print #1, "Dim response_object As Dictionary"
        Print #1, ""
        Print #1, "    exist = False"
        Print #1, ""
        Print #1, "    Set request = New MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "    url = " & vUrl & " & " & Chr(34) & "/" & LCase(vEndpoint) & "/" & Chr(34) & " & " & LCase(vEndpoint) & "Id"
        Print #1, ""
        Print #1, "    request.Open " & Chr(34) & "GET" & Chr(34) & ", url"
        Print #1, "    request.send"
        Print #1, ""
        Print #1, "    Set response_object = modJsonConverter.ParseJson(request.responseText)"
        Print #1, ""
        Print #1, "    If response_object.Item(""status"") = """" Then exist = True"
        Print #1, ""
        Print #1, "    Set request = Nothing"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If
    
End Sub

Public Sub writeAdd()
Dim strLinea As Variant

Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Hace 1 INSERT con los valores de los atributos"
        Print #1, "Public Function add(" & LCase(vEndpoint) & " As " & vClassName & ") As " & vClassName
        Print #1, "Dim request As MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "Dim url As String"
        Print #1, "Dim json_text As String"
        Print #1, ""
        Print #1, "Dim response_object As Dictionary"
        Print #1, ""
        Print #1, "Dim object As New " & vClassName
        Print #1, ""
        Print #1, "    Set add = New " & vClassName
        Print #1, ""
        Print #1, "    Set request = New MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "    url = " & vUrl & " & " & Chr(34) & "/" & LCase(vEndpoint) & "/" & Chr(34)
        Print #1, ""
        Print #1, "    request.Open " & Chr(34) & "POST" & Chr(34) & ", url"
        Print #1, "    request.setRequestHeader ""Content-Type"", ""application/json"""
        Print #1, ""
        Print #1, "    json_text = modJsonConverter.ConvertToJson(" & LCase(vEndpoint) & ".makeParams)"
        Print #1, ""
        Print #1, "    request.send json_text"
        Print #1, ""
        Print #1, "    Set response_object = modJsonConverter.ParseJson(request.responseText)"
        Print #1, ""
        Print #1, "    object.fillobject response_object"
        Print #1, ""
        Print #1, "    If response_object.Item(""status"") = """" Then Set add = object"
        Print #1, ""
        Print #1, "    Set request = Nothing"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If
    
End Sub

Public Sub writeUpdate()
Dim strLinea As Variant

Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Hace 1 UPDATE sólo de los atributos que cambiaron respecto del almacenado en la DB"
        Print #1, "Public Function update(" & LCase(vEndpoint) & " As " & vClassName & ", " & LCase(vEndpoint) & "Id As Long) As " & vClassName
        Print #1, "Dim request As MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "Dim url As String"
        Print #1, "Dim json_text As String"
        Print #1, ""
        Print #1, "Dim response_object As Dictionary"
        Print #1, ""
        Print #1, "Dim object As New " & vClassName
        Print #1, ""
        Print #1, "    set update = " & LCase(vEndpoint)
        Print #1, ""
        Print #1, "    Set request = New MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "    url = " & vUrl & " & " & Chr(34) & "/" & LCase(vEndpoint) & "/" & Chr(34) & " & " & LCase(vEndpoint) & "Id"
        Print #1, ""
        Print #1, "    request.Open " & Chr(34) & "PUT" & Chr(34) & ", url"
        Print #1, "    request.setRequestHeader ""Content-Type"", ""application/json"""
        Print #1, ""
        Print #1, "    json_text = modJsonConverter.ConvertToJson(" & LCase(vEndpoint) & ".makeParams(True))"
        Print #1, ""
        Print #1, "    request.send json_text"
        Print #1, ""
        Print #1, "    Set response_object = modJsonConverter.ParseJson(request.responseText)"
        Print #1, ""
        Print #1, "    object.fillObject response_object"
        Print #1, ""
        Print #1, "    If response_object.Item(""status"") = """" Then Set update = object"
        Print #1, ""
        Print #1, "    Set request = Nothing"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If
    
End Sub

Public Sub writeMakeParams()
Dim strLinea As Variant

Dim objAtributo As New clsAtributo

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Escribe makeParams"
        Print #1, "Public Function makeParams(Optional update As Boolean = False) As Dictionary"
        Print #1, "Dim object As New Dictionary"
        Print #1, ""
        For Each objAtributo In vAtributos
            If objAtributo.tipo = "Date" Then
                Print #1, "    object.add " & Chr(34) & firstLower(objAtributo.atributo) & Chr(34) & ", modConv.date2JSON(" & firstLower(objAtributo.atributo) & ")"
            Else
                Print #1, "    object.add " & Chr(34) & firstLower(objAtributo.atributo) & Chr(34) & ", " & firstLower(objAtributo.atributo)
            End If
        Next
        Print #1, ""
        Print #1, "    Set makeParams = object"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If
    
End Sub

Public Sub writeCollectionAll()

    If Trim(vEndpoint) <> "" Then
        printHeader "Arma una colección con todos los registros"
        Print #1, "Public Function collectionAll() As Collection"
        Print #1, "Dim request As MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "Dim url As String"
        Print #1, ""
        Print #1, "    Set request = New MSXML2.ServerXMLHTTP"
        Print #1, ""
        Print #1, "    url = " & vUrl & " & " & Chr(34) & "/" & LCase(vEndpoint) & "/" & Chr(34)
        Print #1, ""
        Print #1, "    request.Open " & Chr(34) & "GET" & Chr(34) & ", url"
        Print #1, "    request.send"
        Print #1, ""
        Print #1, "    Set collectionAll = collectionJSON(request.responseText)"
        Print #1, ""
        Print #1, "    Set request = Nothing"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
        printHeader "Hace la consulta SQL sobre la DB y arma la colección"
        Print #1, "Public Function collectionJSON(jsontext As String) As Collection"
        Print #1, "Dim elemento As " & vClassName
        Print #1, ""
        Print #1, "Dim registros As Collection"
        Print #1, ""
        Print #1, "Dim json_object As Dictionary"
        Print #1, ""
        Print #1, "    If Left(jsonText, 1) <> " & Chr(34) & "[" & Chr(34) & " Or jsonText = " & Chr(34) & "[]" & Chr(34) & " Then"
        Print #1, "        Set collectionJSON = New Collection"
        Print #1, "        Exit Function"
        Print #1, "    End If"
        Print #1, ""
        Print #1, "    Set registros = New Collection"
        Print #1, ""
        Print #1, "    For Each json_object In modJsonConverter.ParseJson(jsontext)"
        Print #1, "        Set elemento = New " & vClassName
        Print #1, "        elemento.fillObject json_object"
        Print #1, ""
        Print #1, "        If Not modCollection.collectionExistElement(registros, elemento.key) Then registros.add elemento, elemento.key"
        Print #1, "    Next"
        Print #1, ""
        Print #1, "    Set collectionJSON = registros"
        Print #1, ""
        Print #1, "End Function"
        Print #1, ""
    End If

End Sub

Public Sub writeFillCombo()
Dim objAtributo As New clsAtributo

Dim strLinea As Variant

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Llena una ComboBox con todos los datos de la tabla"
        Print #1, "Public Sub fillCombo(pCombo As ComboBox)"
        Print #1, ""
        Print #1, "    fillComboCollection pCombo, collectionAll()"
        Print #1, ""
        Print #1, "End Sub"
        Print #1, ""
        printHeader "Llena la ComboBox efectivamente"
        Print #1, "Private Sub fillComboCollection(pCombo As ComboBox, pCollection As Collection)"
        Print #1, "Dim elemento As " & vClassName
        Print #1, ""
        Print #1, "    pCombo.Clear"
        Print #1, ""
        Print #1, "    For Each elemento In pCollection"
        Print #1, "        pCombo.AddItem elemento.comboText"
        Print #1, "        pCombo.ItemData(pCombo.NewIndex) = elemento.comboKey"
        Print #1, "    Next"
        Print #1, ""
        Print #1, "    If pCombo.ListCount > 0 Then pCombo.ListIndex = 0"
        Print #1, ""
        Print #1, "End Sub"
        Print #1, ""
    End If

End Sub

Public Sub writeFillList()
Dim objAtributo As New clsAtributo

Dim strLinea As Variant

    If objAtributo.existPrimary(vAtributos) And Trim(vEndpoint) <> "" Then
        printHeader "Llena una ListBox con todos los datos de la tabla"
        Print #1, "Public Sub fillList(pList As ListBox)"
        Print #1, ""
        Print #1, "    fillListCollection pList, collectionAll()"
        Print #1, ""
        Print #1, "End Sub"
        Print #1, ""
        printHeader "Llena la ListBox efectivamente"
        Print #1, "Private Sub fillListCollection(pList As ListBox, pCollection As Collection)"
        Print #1, "Dim elemento As " & vClassName
        Print #1, ""
        Print #1, "    pList.Clear"
        Print #1, ""
        Print #1, "    For Each elemento In pCollection"
        Print #1, "        pList.AddItem elemento.listText"
        Print #1, "        pList.ItemData(pList.NewIndex) = elemento.listKey"
        Print #1, "    Next"
        Print #1, ""
        Print #1, "End Sub"
        Print #1, ""
    End If

End Sub

Private Function printHeader(description As String)

    Print #1, "' * * * * * * * * * * * * * * * * * * * * *"
    Print #1, "' Author: Ing. Daniel Quinteros"
    Print #1, "' Created: " & Date
    Print #1, "' Description: " & description
    Print #1, "' Log: "
    Print #1, "' * * * * * * * * * * * * * * * * * * * * *"

End Function
