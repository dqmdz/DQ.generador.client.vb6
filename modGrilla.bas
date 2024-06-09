Attribute VB_Name = "modGrilla"
Option Explicit

Public Sub makeGrid(pGrid As MSFlexGrid, pTitulos As Variant, pAnchos As Variant, pFixedCols As Integer, pFixedRows As Integer, pMode As Integer)
Dim intColumnas As Integer
Dim intColumna As Integer
    
    intColumnas = UBound(pTitulos) + 1
    pGrid.Clear
    pGrid.SelectionMode = pMode
    pGrid.ScrollBars = flexScrollBarVertical
    pGrid.Rows = 2
    pGrid.Cols = intColumnas
    pGrid.FixedCols = pFixedCols
    pGrid.FixedRows = pFixedRows
    pGrid.Rows = 1
    For intColumna = 0 To intColumnas - 1
        pGrid.TextMatrix(0, intColumna) = pTitulos(intColumna)
        pGrid.ColWidth(intColumna) = pAnchos(intColumna)
    Next intColumna

End Sub

Public Sub setCheckGrid(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer, pValue As Boolean)
    
    pGrid.Row = pRow
    pGrid.Col = pCol
    pGrid.CellFontName = "Wingdings"
    pGrid.CellFontSize = 14
    pGrid.CellAlignment = 4
    pGrid.Text = IIf(pValue, Chr(254), Chr(113))

End Sub

Public Function getCheckGrid(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer) As Boolean
    
    If pGrid.TextMatrix(pRow, pCol) = "" Then
        getCheckGrid = False
        Exit Function
    End If
    
    getCheckGrid = IIf(Asc(pGrid.TextMatrix(pRow, pCol)) = 254, True, False)

End Function

Public Function calcTopGrid(pGrid As MSFlexGrid) As Long
    
    calcTopGrid = pGrid.Top + pGrid.CellTop

End Function

Public Function calcLeftGrid(pGrid As MSFlexGrid) As Long
    
    calcLeftGrid = pGrid.Left + pGrid.CellLeft

End Function

Public Sub setColorGrid(pGrid As MSFlexGrid, pRow As Integer, pCol As Integer, pColor As Long)
    
    pGrid.Row = pRow
    pGrid.Col = pCol
    pGrid.CellBackColor = pColor

End Sub
