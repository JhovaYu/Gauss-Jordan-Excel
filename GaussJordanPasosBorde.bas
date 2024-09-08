Attribute VB_Name = "GaussJordanPasosBorde"
Sub GaussJordanConPasosConBordes4()
    On Error GoTo ErrorHandler
    
    
    If Selection.Rows.Count < 2 Or Selection.Columns.Count < 2 Then
        MsgBox "Por favor, selecciona una matriz de al menos 2x2."
        Exit Sub
    End If
    
    
    Dim matrizOriginal() As Double
    Dim matrizTransformada() As Double
    Dim numRows As Integer
    Dim numCols As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim factor As Double
    Dim paso As Integer
    Dim startCol As Integer
    Dim lastCol As Integer
    Dim startRow As Integer

    
    numRows = Selection.Rows.Count
    numCols = Selection.Columns.Count

    
    ReDim matrizOriginal(1 To numRows, 1 To numCols)
    ReDim matrizTransformada(1 To numRows, 1 To numCols)
    
    
    For i = 1 To numRows
        For j = 1 To numCols
            matrizOriginal(i, j) = Selection.Cells(i, j).Value
            matrizTransformada(i, j) = matrizOriginal(i, j)
        Next j
    Next i

    
    paso = 0
    startCol = numCols + 2
    startRow = 1

    
    For i = 1 To numRows
        paso = paso + 1

        
        For j = 1 To numRows
            For k = 1 To numCols
                Selection.Cells(startRow + j - 1, startCol + (paso - 1) * (numCols + 1) + k - 1).Value = matrizTransformada(j, k)
            Next k
        Next j
        
        
        With Selection.Range(Selection.Cells(startRow - 12, startCol + (paso - 1) * (numCols + 1)), _
                             Selection.Cells(startRow + numRows - 13, startCol + (paso - 1) * (numCols + 1) + numCols - 1)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
        
        With Selection.Range(Selection.Cells(startRow - 12, startCol + (paso - 1) * (numCols + 1) + numCols - 1), _
                             Selection.Cells(startRow + numRows - 13, startCol + (paso - 1) * (numCols + 1) + numCols - 1)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        
        
        If matrizTransformada(i, i) = 0 Then
            
            For k = i + 1 To numRows
                If matrizTransformada(k, i) <> 0 Then
                    For j = 1 To numCols
                        factor = matrizTransformada(i, j)
                        matrizTransformada(i, j) = matrizTransformada(k, j)
                        matrizTransformada(k, j) = factor
                    Next j
                    Exit For
                End If
            Next k
            If matrizTransformada(i, i) = 0 Then
                MsgBox "El sistema no tiene solución única o es inconsistente."
                Exit Sub
            End If
        End If
        
        factor = matrizTransformada(i, i)
        For j = 1 To numCols
            matrizTransformada(i, j) = matrizTransformada(i, j) / factor
        Next j
        
        For k = 1 To numRows
            If k <> i Then
                factor = matrizTransformada(k, i)
                For j = 1 To numCols
                    matrizTransformada(k, j) = matrizTransformada(k, j) - factor * matrizTransformada(i, j)
                Next j
            End If
        Next k
    Next i

    paso = paso + 1
    For i = 1 To numRows
        For j = 1 To numCols
            Selection.Cells(startRow + i - 1, startCol + (paso - 1) * (numCols + 1) + j - 1).Value = matrizTransformada(i, j)
        Next j
    Next i
    
    With Selection.Range(Selection.Cells(startRow - 12, startCol + (paso - 1) * (numCols + 1)), _
                         Selection.Cells(startRow + numRows - 13, startCol + (paso - 1) * (numCols + 1) + numCols - 1)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Range(Selection.Cells(startRow - 12, startCol + (paso - 1) * (numCols + 1) + numCols - 1), _
                         Selection.Cells(startRow + numRows - 13, startCol + (paso - 1) * (numCols + 1) + numCols - 1)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .TintAndShade = 0
        .Weight = xlMedium
    End With

    MsgBox "El método de Gauss-Jordan ha sido aplicado exitosamente, mostrando cada paso de la transformación con bordes."
    Exit Sub

ErrorHandler:
    MsgBox "Se ha producido un error: " & Err.Description
End Sub


