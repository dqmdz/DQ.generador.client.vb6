VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub fillCombo(pCombo As ComboBox)
Dim varTipos As Variant

Dim intCiclo As Integer

    varTipos = Array("Date", "Integer", "Long", "String", "Double", "Variant", "Time", "Currency")
    
    pCombo.Clear
    
    For intCiclo = LBound(varTipos) To UBound(varTipos)
        pCombo.AddItem varTipos(intCiclo)
    Next intCiclo
    
End Sub

Public Function initialValue(pAtributo As clsAtributo) As String

    If pAtributo.nullable Then
        initialValue = "Null"
        Exit Function
    End If
    
    Select Case pAtributo.tipo
        Case "Date"
            initialValue = "Date"
        Case "Time"
            initialValue = "CDate(" & Chr(34) & "00:00:00" & Chr(34) & ")"
        Case "Integer", "Long", "Double", "Currency"
            initialValue = "0"
        Case "String"
            initialValue = Chr(34) & Chr(34)
    End Select
    
End Function

