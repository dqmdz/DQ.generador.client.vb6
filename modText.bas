Attribute VB_Name = "modText"
Option Explicit

Public Sub markText(pText As TextBox)

    pText.SelStart = 0
    pText.SelLength = Len(pText.Text)
    
End Sub

Public Function firstUpper(pString As String)

    firstUpper = UCase(Left(pString, 1)) & Mid(pString, 2)
    
End Function

Public Function firstLower(pString As String)

    firstLower = LCase(Left(pString, 1)) & Mid(pString, 2)
    
End Function

