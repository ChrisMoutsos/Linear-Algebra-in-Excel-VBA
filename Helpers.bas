Sub matrixAddRow( _
    ByRef matrix, _
    m, n, _
    sourceIndex, targetIndex, _
    multiplier _
)
    ' Purpose: Adds multiplier*(sourceIndex-th row) to targetIndex-th row
    ' Inputs: matrix - the matrix
    '             m - number of rows in matrix
    '             n - number of columns in matrix
    '             sourceIndex - row index to be added
    '             targetIndex - row index to change
    '            multiplier - how much to add sourceIndex-th row to targetIndex-th row
    Dim j As Integer
    For j = 1 To n
        matrix(targetIndex, j) = matrix(targetIndex, j) + multiplier * matrix(sourceIndex, j)
    Next j
End Sub

Sub matrixMultiplyRow( _
    ByRef matrix, _
    m, n, _
    targetIndex, _
    multiplier _
)
    ' Purpose: Multiplies multiplier*(targetIndex-th row)
    ' Inputs: matrix - the matrix
    '             m - number of rows in matrix
    '             n - number of columns in matrix
    '             targetIndex - row index to change
    '            multiplier - how much to multiply by
    Dim j As Integer
    For j = 1 To n
        matrix(targetIndex, j) = matrix(targetIndex, j) * multiplier
    Next j
End Sub

Sub matrixSwapRow( _
    ByRef matrix, _
    m, n, _
    indexOne, indexTwo _
)
    ' Purpose: Swaps rows indexOne and indexTwo of matrix
    ' Inputs: matrix - the matrix
    '             m - number of rows in matrix
    '             n - number of columns in matrix
    '             indexOne - first row index to swap
    '             indexTwo - second row index to swap
    Dim temp As Long
    Dim j As Integer
    For j = 1 To n
         temp = matrix(indexOne, j)
         matrix(indexOne, j) = matrix(indexTwo, j)
         matrix(indexTwo, j) = temp
    Next j
End Sub

Sub fillAreaByMatrix( _
    ByRef matrix, _
    m, n, _
    area As range _
)
    ' Purpose: Write matrix to a given area (starting from top-left cell)
    ' Inputs: matrix - the matrix with values to write
    '             m - number of rows in matrix
    '             n - number of columns in matrix
    '             area - the area to write to
    Dim i, j As Integer
    For i = 1 To m
        For j = 1 To n
            Cells(area.Row + i - 1, area.Column + j - 1) = matrix(i, j)
        Next j
    Next i
End Sub

Function fillMatrixByArea( _
    ByRef matrix, _
    area, _
    Optional asTranpose As Boolean = False _
)
    ' Purpose: Initialize a given matrix with values from an area.
    '               Area and matrix must be the same size.
    ' Inputs: matrix - an mxn matrix
    '             area - an mxn area
    '             asTranpose - whether to read into matrix as the tranpose
    '                                   (give an mxn matrix and and nxm selection)
    Dim rangeCell As range
    Dim relativeRow, relativeColumn As Integer
    For Each rangeCell In area.Cells
        relativeRow = rangeCell.Row - area.Row
        relativeColumn = rangeCell.Column - area.Column
        If asTranpose Then
            matrix(relativeColumn + 1, relativeRow + 1) = rangeCell.Value2
        Else
            matrix(relativeRow + 1, relativeColumn + 1) = rangeCell.Value2
        End If
    Next rangeCell
    fillMatrixByArea = matrix
End Function

Function createMatrix(m, n)
    ' Purpose: Return an mxn matrix.
    ' Inputs: m - number of rows
    '             n - number of columns
    Dim matrix() As Double
    ReDim matrix(1 To m, 1 To n) As Double
    createMatrix = matrix
End Function

Sub matrixRowReduce( _
    ByRef matrix, _
    m, n, _
    Optional withIdentity = False _
)
    ' Purpose: Converts matrix into row echelon form.
    ' Inputs: matrix - matrix to convert
    '             m - number of rows
    '             n - number of columns
    
    Dim adjustedN As Integer
    If withIdentity = True Then
        adjustedN = n * 2
    Else
        adjustedN = n
    End If

    Dim startRow, startColumn As Integer
    startRow = 1
    startColumn = 1
    Do While startRow <= m And startColumn <= n
        ' Step 1. Find the first (from the left) nonzero column.
        Dim allZeroes As Boolean
        Dim firstNonZeroColumn, firstNonZeroRow As Integer
        firstNonZeroColumn = startColumn
        firstNonZeroRow = startRow
        For j = firstNonZeroColumn To n
            allZeroes = True
            For i = startRow To m
                If matrix(i, j) <> 0 Then
                    allZeroes = False
                    firstNonZeroRow = i
                    Exit For
                End If
            Next i
            If allZeroes = False Then
                firstNonZeroColumn = j
                Exit For
            End If
        Next j
        
        ' Step 2. If the first nonzero column is the jth column, use row
        '            operations to make matrix(1, firstNonZeroColumn) <> 0.
        '            The entry matrix(1, firstNonZeroColumn) will be a pivot.
        If matrix(startRow, firstNonZeroColumn) = 0 Then
            Dim swapRow As Integer
            For i = (startRow + 1) To m
                If matrix(i, firstNonZeroColumn) <> 0 Then
                    Helpers.matrixSwapRow _
                        matrix:=matrix, _
                        m:=m, _
                        n:=adjustedN, _
                        indexOne:=startRow, _
                        indexTwo:=i
                    Exit For
                End If
            Next i
        End If
        
        ' Step 3: Use row operations to make all entries in the
        '            column below the pivot equal to 0, i.e. make
        '            matrix(2, firstNonZeroColumn) =
        '            matrix(3, firstNonZeroColumn) = ... = 0
        Dim mult As Double
        For i = (startRow + 1) To m
            If matrix(i, firstNonZeroColumn) <> 0 Then
                mult = (-1 * _
                    (matrix(i, firstNonZeroColumn) / _
                    matrix(startRow, firstNonZeroColumn)) _
                )
                Helpers.matrixAddRow _
                    matrix:=matrix, _
                    m:=m, _
                    n:=adjusteN, _
                    sourceIndex:=startRow, _
                    targetIndex:=i, _
                    multiplier:=mult
            End If
        Next i
        
        ' Step 4: Let "new matrix" be the (m-1)x(n-1) matrix obtained
        '             from "old matrix" by deleting the first row
        '             and the first firstNonZeroColumn columns.
        startRow = startRow + 1
        startColumn = startColumn + firstNonZeroColumn
    Loop
End Sub

Sub matrixReducePivots( _
    ByRef matrix, _
    m, n, _
    Optional withIdentity = False _
)
    ' Purpose: Converts row echeleon form matrix into
    '               reduced row echelon form.
    ' Inputs: matrix - matrix to convert, in row echelon form
    '             m - number of rows
    '             n - number of columns
    
    Dim adjustedN As Integer
    If withIdentity = True Then
        adjustedN = n * 2
    Else
        adjustedN = n
    End If
    
    Dim startRow, startColumn As Integer
    startRow = m
    startColumn = n
    Do While startRow >= 1 And startColumn >= 1
        ' Step 1: Use row operations to make all pivots equal to one
        Dim i, j, pivotCol As Integer
        Dim mult As Double
        For i = 1 To m
            pivotCol = 0
            For j = 1 To n
                If matrix(i, j) <> 0 Then
                    mult = 1 / matrix(i, j)
                    Helpers.matrixMultiplyRow _
                        matrix:=matrix, _
                        m:=m, _
                        n:=adjustedN, _
                        targetIndex:=i, _
                        multiplier:=mult
                    Exit For
                End If
            Next j
        Next i
        
        ' Step 2: Identify the lowest pivot, the pivot closest
        ' to the bottom right corner of the matrix.
        Dim lowestPivotRow, lowestPivotColumn As Integer
        lowestPivotRow = 0
        lowestPivotColumn = 0
        For i = startRow To 1 Step -1
            For j = startColumn To 1 Step -1
                If matrix(i, j) <> 0 Then
                    lowestPivotRow = i
                    lowestPivotColumn = j
                    Exit For
                End If
            Next j
            If lowestPivotRow And lowestPivotColumn Then
                Exit For
            End If
        Next i
        
        If lowestPivotRow And lowestPivotColumn Then
            For i = (lowestPivotRow - 1) To 1 Step -1
                If matrix(i, lowestPivotColumn) <> 0 Then
                    mult = (-1 * _
                        (matrix(i, lowestPivotColumn) / _
                        matrix(lowestPivotRow, lowestPivotColumn)) _
                    )
                    Helpers.matrixAddRow _
                        matrix:=matrix, _
                        m:=m, _
                        n:=adjustedN, _
                        sourceIndex:=lowestPivotRow, _
                        targetIndex:=i, _
                        multiplier:=mult
                End If
            Next i
        End If
        
        ' Step 4. Let "new matrix" be the (i-1)x(j-1) matrix
        ' consisting of the first (i-1) rows and (j-1) columns of "old matrix"
        startRow = startRow - 1
        startColumn = startColumn - 1
    Loop
 End Sub
 
 Function invertMatrix( _
    ByRef matrix, _
    m, n _
)
    ' Purpose: Inverts matrix (needs to be square).
    ' Inputs: matrix - matrix to convert
    '             m - number of rows
    '             n - number of columns
    
    Dim i, j As Integer
    matrixAndIdentity = Helpers.createMatrix(m, 2 * m)
    For i = 1 To m
        For j = 1 To n
            matrixAndIdentity(i, j) = matrix(i, j)
        Next j
    Next i
    For i = 1 To m
        For j = (m + 1) To 2 * m
            If i = (j - m) Then
                matrixAndIdentity(i, j) = 1
            Else
                matrixAndIdentity(i, j) = 0
            End If
        Next j
    Next i
    
    ' Row reduce matrix A to row echelon form REF(A)
    Helpers.matrixRowReduce _
        matrix:=matrixAndIdentity, _
        m:=m, _
        n:=n, _
        withIdentity:=True
        
    ' Fully row reduce REF(A) to reduced row echelon form RREF(A)
    Helpers.matrixReducePivots _
        matrix:=matrixAndIdentity, _
        m:=m, _
        n:=n, _
        withIdentity:=True
        
    ' Grab A^-1 (what the identity matrix transformed into)
    invertedMatrix = Helpers.createMatrix(m, n)
    For i = 1 To m
        For j = 1 To n
            invertedMatrix(i, j) = matrixAndIdentity(i, j + n)
        Next j
    Next i
    invertMatrix = invertedMatrix
End Function
    
    


