VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub writeHeader(pFile As Integer, pClassName As String)
    
    Print #pFile, "VERSION 1.0 CLASS"
    Print #pFile, "BEGIN"
    Print #pFile, "  MultiUse = -1  'True"
    Print #pFile, "  Persistable = 0  'NotPersistable"
    Print #pFile, "  DataBindingBehavior = 0  'vbNone"
    Print #pFile, "  DataSourceBehavior  = 0   'vbNone"
    Print #pFile, "  MTSTransactionMode  = 0   'NotAnMTSObject"
    Print #pFile, "END"
    Print #pFile, "Attribute VB_Name = " & Chr(34) & pClassName & Chr(34)
    Print #pFile, "Attribute VB_GlobalNameSpace = False"
    Print #pFile, "Attribute VB_Creatable = True"
    Print #pFile, "Attribute VB_PredeclaredId = False"
    Print #pFile, "Attribute VB_Exposed = False"
    Print #pFile, "Option Explicit"
    Print #pFile, ""

End Sub
