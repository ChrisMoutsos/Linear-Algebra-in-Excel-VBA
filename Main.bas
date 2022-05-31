Sub btnGenerateRandom_Click()
    ' Purpose: Fill in the selections with random integers.
    
    Dim randomNumber, areaIndex As Integer
    Dim rangeCell As range
    
    For areaIndex = 1 To Selection.Areas.Count
        For Each rangeCell In Selection.Areas(areaIndex).Cells
            randomNumber = Application.WorksheetFunction.RandBetween(-100, 100)
            range(rangeCell.address) = randomNumber
        Next rangeCell
    Next areaIndex
End Sub

Sub btnGenerateIdentity_Click()
    ' Purpose: Fill in the selections (must be square) with the identity matrix.

    Dim areaIndex, relativeRow, relativeColumn As Integer
    Dim rangeCell As range
    
    For areaIndex = 1 To Selection.Areas.Count
        
        ' Require a square matrix
        If Selection.Areas(areaIndex).Columns.Count <> Selection.Areas(areaIndex).Rows.Count Then
            MsgBox "Please select a square area (mxm)"
        Else
            For Each rangeCell In Selection.Areas(areaIndex).Cells
                ' 1s on diagonal, 0s everywhere else
                relativeRow = rangeCell.Row - Selection.Areas(areaIndex).Row
                relativeColumn = rangeCell.Column - Selection.Areas(areaIndex).Column
                If relativeRow = relativeColumn Then
                    value = 1
                Else
                    value = 0
                End If
                Cells(rangeCell.Row, rangeCell.Column) = value
            Next rangeCell
        End If
    Next areaIndex
End Sub

Sub btnMultiplyMatrices_Click()
    ' Purpose: Given three ordered selections A, B, C,
    '               multiply A and B as matrices and put
    '               the result into the cells starting at top-left of C.

    Dim areaIndex As Integer

    ' We'll be doing AB=C
    Dim matrixARowCount, matrixAColumnCount As Integer
    Dim matrixBRowCount, matrixBColumnCount As Integer
    Dim matrixCRowCount, matrixCColumnCount As Integer
    
    ' Require 3 areas selected
    If Selection.Areas.Count <> 3 Then
        MsgBox "Please select three areas (matrix A, matrix B, and top-left of result matrix destination)"
        Exit Sub
    End If
    
    ' Grab sizes of areas
    matrixARowCount = Selection.Areas(1).Rows.Count
    matrixAColumnCount = Selection.Areas(1).Columns.Count
    matrixBRowCount = Selection.Areas(2).Rows.Count
    matrixBColumnCount = Selection.Areas(2).Columns.Count
    matrixCRowCount = matrixARowCount
    matrixCColumnCount = matrixBColumnCount
        
    ' Require sizes A as (mxn) and B as (nxp)
    If matrixAColumnCount <> matrixBRowCount Then
        MsgBox "Column count of matrix A must match row count of matrix B"
        Exit Sub
    End If
    
    ' Now we are ready to start the fun
    matrixA = Helpers.createMatrix(matrixARowCount, matrixAColumnCount)
    matrixB = Helpers.createMatrix(matrixBRowCount, matrixBColumnCount)
    matrixC = Helpers.createMatrix(matrixCRowCount, matrixCColumnCount)
    
    matrixA = Helpers.fillMatrixByArea(matrixA, Selection.Areas(1))
    matrixB = Helpers.fillMatrixByArea(matrixB, Selection.Areas(2))
    
    Dim i, j, k, value As Integer
    For i = 1 To matrixCRowCount
        For j = 1 To matrixCColumnCount
            value = 0
            For k = 1 To matrixAColumnCount
                value = value + (matrixA(i, k) * matrixB(k, j))
            Next k
            matrixC(i, j) = value
        Next j
    Next i
    
    ' Write the result to the sheet
    Helpers.fillAreaByMatrix _
        matrix:=matrixC, _
        m:=matrixCRowCount, _
        n:=matrixCColumnCount, _
        area:=Selection.Areas(3)
    
End Sub

Sub btnTranspose_Click()
    ' Purpose: Transpose the first selected matrix into the
    '                cells starting at the top-left of the second selection.

    Dim relativeRow, relativeColumn, _
    matrixRowCount, matrixColumnCount As Integer
    Dim rangeCell, matrixArea, resultArea As range
    
    ' Require 2 areas selected
    If Selection.Areas.Count <> 2 Then
        MsgBox "Please select two areas (matrix A and top-left of A tranpose destination)"
        Exit Sub
    End If
    
    ' Create our transpose matrix from first selection
    Set matrixArea = Selection.Areas(1)
    matrixRowCount = matrixArea.Rows.Count
    matrixColumnCount = matrixArea.Columns.Count
    matrixTranspose = Helpers.createMatrix(matrixColumnCount, matrixRowCount)
    matrixTranspose = Helpers.fillMatrixByArea(matrixTranspose, matrixArea, True)
    
    ' Write the transpose result back to the second selection
    Set resultArea = Selection.Areas(2)
    Helpers.fillAreaByMatrix _
        matrix:=matrixTranspose, _
        m:=matrixColumnCount, _
        n:=matrixRowCount, _
        area:=resultArea
End Sub

Sub btnRowReduce_Click()
    ' Purpose: Row reduce the matrix (first selection)
    '               and put the row echelon form into the cells
    '               starting at the top-left of the second selection.

    Dim matrixRowCount, matrixColumnCount, _
    i, j As Integer
    Dim matrixArea, resultArea As range
    
    ' Require 2 areas selected
    If Selection.Areas.Count <> 2 Then
        MsgBox "Please select two areas (matrix A and top-left of REF(A) destination)"
        Exit Sub
    End If
    
    ' Grab our source matrix from sheet
    Set matrixArea = Selection.Areas(1)
    matrixRowCount = CInt(matrixArea.Rows.Count)
    matrixColumnCount = CInt(matrixArea.Columns.Count)
    matrix = Helpers.createMatrix(matrixRowCount, matrixColumnCount)
    matrix = Helpers.fillMatrixByArea(matrix, matrixArea)
    
     ' Row reduce matrix A to row echelon form REF(A)
    Helpers.matrixRowReduce _
        matrix:=matrix, _
        m:=matrixRowCount, _
        n:=matrixColumnCount
    
    ' Write the result, REF(A), back to the second selection
    Set resultArea = Selection.Areas(2)
    Helpers.fillAreaByMatrix _
        matrix:=matrix, _
        m:=matrixRowCount, _
        n:=matrixColumnCount, _
        area:=resultArea
    
End Sub

Sub btnFullRowReduce_Click()
    ' Purpose: Fully row reduce the matrix (first selection)
    '               and put the reduced row echelon form into the cells
    '               starting at the top-left of the second selection.

    Dim matrixRowCount, matrixColumnCount, _
    i, j As Integer
    Dim matrixArea, resultArea As range
    
    ' Require 2 areas selected
    If Selection.Areas.Count <> 2 Then
        MsgBox "Please select two areas (matrix A and top-left of RREF(A) destination)"
        Exit Sub
    End If
    
    ' Grab our source matrix from sheet
    Set matrixArea = Selection.Areas(1)
    matrixRowCount = matrixArea.Rows.Count
    matrixColumnCount = matrixArea.Columns.Count
    matrix = Helpers.createMatrix(matrixRowCount, matrixColumnCount)
    matrix = Helpers.fillMatrixByArea(matrix, matrixArea)
    
     ' Row reduce matrix A to row echelon form REF(A)
    Helpers.matrixRowReduce _
        matrix:=matrix, _
        m:=matrixRowCount, _
        n:=matrixColumnCount
        
    ' Fully row reduce REF(A) to reduced row echelon form RREF(A)
    Helpers.matrixReducePivots _
        matrix:=matrix, _
        m:=matrixRowCount, _
        n:=matrixColumnCount
    
    ' Write the result, RREF(A), back to the second selection
    Set resultArea = Selection.Areas(2)
    Helpers.fillAreaByMatrix _
        matrix:=matrix, _
        m:=matrixRowCount, _
        n:=matrixColumnCount, _
        area:=resultArea
End Sub

Sub btnInvertMatrix_Click()
    Dim resultArea As range
    ' Require 2 areas selected
    If Selection.Areas.Count <> 2 Then
        MsgBox "Please select two areas (matrix A and top-left of A^-1 destination)"
        Exit Sub
    End If
    
    If Selection.Areas(1).Columns.Count <> Selection.Areas(1).Rows.Count Then
            MsgBox "Please select a square matrix A (mxm)"
            Exit Sub
    End If
    
    ' Grab our source matrix from sheet
    Set matrixArea = Selection.Areas(1)
    matrixRowCount = matrixArea.Rows.Count
    matrixColumnCount = matrixArea.Columns.Count
    matrix = Helpers.createMatrix(matrixRowCount, matrixColumnCount)
    matrix = Helpers.fillMatrixByArea(matrix, matrixArea)
    
    ' Invert matrix
    invertedMatrix = Helpers.invertMatrix( _
        matrix:=matrix, _
        m:=matrixRowCount, _
        n:=matrixColumnCount _
    )
        
    ' Write the result, A^-1, back to the second selection
    Set resultArea = Selection.Areas(2)
    Helpers.fillAreaByMatrix _
        matrix:=invertedMatrix, _
        m:=matrixRowCount, _
        n:=matrixColumnCount, _
        area:=resultArea
End Sub

