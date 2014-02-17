Attribute VB_Name = "Adj"
'Вспомогательная библиотечка

Public Function selectionCheck(ByVal square As Boolean) As Boolean ' Проверка на правильность выделения исходной матрицы
    selectionCheck = True
    If Application.selection.Cells.count <= 1 Then
        MsgBox ("Матрица не выделена!")
        selectionCheck = False
        Exit Function
    End If
    If square = True Then
        If ((Application.selection.rows.count + 1) <> Application.selection.Columns.count) Then
            MsgBox ("Матрица не квадратная!!!!!")
            selectionCheck = False
            Exit Function
        End If
    End If
End Function

Public Sub Randomize() 'Заполнить выделение случайными данными
    For Each Cell In selection
        Cell.Value = Int(200 * Rnd - 100)
    Next
End Sub

Public Sub RndAndCalcGauss() 'Кнопка для Гаусса - Случайные данные и пересчитать
    Dim NowSelection As Range
    Set NowSelection = Application.selection
    Call Randomize
    Call GaussSolve.GaussSolve
    NowSelection.Select
'Очищение памяти
Set NowSelection = Nothing
End Sub

Public Sub SubCalculations() 'Флажок для перебора, отображение промежуточных расчетов
    BazisSolve.isSubCalc = Not BazisSolve.isSubCalc
End Sub

