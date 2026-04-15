Attribute VB_Name = "MdlSpread"
Option Explicit

Function grdSetText(ByRef ngrid As vaSpread, ByVal nCol As Long, _
                      ByVal nRow As Long, ByVal nValue As Variant)
    On Error GoTo grdSetText_Error

    ngrid.Row = nRow
    ngrid.Col = nCol
    ngrid.text = nValue

    Exit Function
grdSetText_Error:
End Function

Function grdInsertRow(ByVal ngrid As vaSpread, nRow As Long)

    With ngrid
        
        .MaxRows = .MaxRows + 1
        .Row = nRow
        .Action = 7

    End With

End Function

Function grdAddRow(ByVal ngrid As vaSpread)

    With ngrid
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
    End With

End Function

Function grdCellTypeStatic(ByVal ngrid As vaSpread, ByVal nCol As Long, ByVal nRow As Long, ByVal Haling As Integer)

    With ngrid
        .Row = nRow
        .Col = nCol
        .CellType = CellTypeStaticText
        .TypeHAlign = IIf(Haling = 0, TypeHAlignLeft, IIf(Haling = 1, TypeHAlignRight, TypeHAlignCenter))
    End With
    
End Function

Function grdCellTypeCkeckBox(ByVal ngrid As vaSpread, ByVal nCol As Long, ByVal nRow As Long, ByVal Haling As Integer)

    With ngrid
        .Row = nRow
        .Col = nCol
        .CellType = CellTypeCheckBox
        .TypeHAlign = IIf(Haling = 0, TypeHAlignLeft, IIf(Haling = 1, TypeHAlignRight, TypeHAlignCenter))
    End With
    
End Function

Function grdCellTypeEdit(ByVal ngrid As vaSpread, ByVal nCol As Long, ByVal nRow As Long, ByVal Haling As Integer, ByVal maxLen As Long)

    With ngrid
        .Row = nRow
        .Col = nCol
        .CellType = CellTypeEdit
        .TypeMaxEditLen = maxLen
        .TypeHAlign = IIf(Haling = 0, TypeHAlignLeft, IIf(Haling = 1, TypeHAlignRight, TypeHAlignCenter))
    End With
    
End Function


Function grdCellTypeCurrency(ByVal ngrid As vaSpread, ByVal nCol As Long, ByVal nRow As Long, ByVal Haling As Integer)

    With ngrid
         .Row = nRow
         .Col = nCol
         .CellType = CellTypeCurrency
         .TypeCurrencyDecimal = "."
         .TypeCurrencyDecPlaces = 0
         .TypeCurrencyNegStyle = TypeCurrencyNegStyle1
         .TypeCurrencyPosStyle = TypeCurrencyPosStyle1
         .TypeCurrencySeparator = ","
         .TypeCurrencyShowSep = False
         .TypeCurrencyShowSymbol = False
         .TypeCurrencySymbol = "$"
        .TypeHAlign = IIf(Haling = 0, TypeHAlignLeft, TypeHAlignRight)
    End With
    
End Function

Function grdRowForeColor(ByRef Grid As vaSpread, ByVal Row As Long, ByVal Cor As Variant)

    With Grid
        .Row = Row
        .Row2 = Row
        .Col = -1
        .ForeColor = Cor
    End With

End Function

Function grdRowColForeColor(ByRef Grid As vaSpread, ByVal Row1 As Long, ByVal Row2 As Long, ByVal Col1 As Long, ByVal Col2 As Long, ByVal Cor As Variant)
    
    With Grid
        .BlockMode = True
        
        .Row = Row1
        .Row2 = Row2
        
        .Col = Col1
        .Col2 = Col2
        
        .ForeColor = Cor
        
        .BlockMode = False
    End With

End Function
