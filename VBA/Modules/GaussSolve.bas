Attribute VB_Name = "GaussSolve"
Private gauss As clsGauss
Private rows_index() As Integer
Private cols_index() As Integer


Public Sub GaussSolve() 'Для решения квадратной матрицы методом Гаусса, вывод результата

    On Error GoTo exitGaussSolve:
    If Adj.selectionCheck(True) = False Then
        GoTo exitGaussSolve:
    End If
    Set gauss = New clsGauss
    Call gauss.Initialize
    gauss.max_iterration = 3
    Call init_index
    If gauss.CalculateMatrix(rows_index(), cols_index()) = False Then
        Exit Sub
    End If
    Call gauss.createTables(Application.selection.Offset(gauss.rows_count + 3, 0), rows_index(), cols_index())

exitGaussSolve:
Set gauss = Nothing
Erase rows_index()
Erase cols_index()
End Sub

Private Sub init_index()
    Call gauss.init_rowIndeces(rows_index())
    Call gauss.init_colIndeces(cols_index())
End Sub

