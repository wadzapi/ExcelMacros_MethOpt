VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBazis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private comb As clsCombinations 'Экземпляр класса Сочетание
Private gauss As clsGauss 'Экземпляр класс Гаусс
Private col_indeces() As Integer 'Массив индексов столбцов
Private row_indeces() As Integer 'Массив индексов строк
Private bazis() As Double 'Сводная матрица полученных базисов
Private bazis_titles() As String 'Массив с подписями базисов
Private baz_Count As Integer 'Число строк в массиве названий базисов
Public CreateSubTables As Boolean 'Индикатор создания подтаблиц, полученных в результате комбинаций
Private SubTableRange As Range 'Диапазон для построения подтаблицы
Private RefSolutions() As Double 'Массив опорных решений
Private RefSolutionsTitles() As String 'Массив подписей опорных решений
Private RefSolutionsCount As Integer 'Число опорных решений


Public Sub CalculateBazis() 'Вычисление базисных решений
    If CreateSubTables = True Then 'Сохраняет диапазон выделения для построения подтаблиц
        Set SubTableRange = Application.selection
    End If
    Do
        Call comb.getColsIndex(col_indeces()) 'Получение новых сочетаний индексов столбцов матрицы
        Call gauss.init_rowIndeces(row_indeces()) 'Новые индексы столбцов
        If gauss.CalculateMatrix(row_indeces(), col_indeces()) = False Then 'Решение матрицы по Гауссу
            Exit Sub
        End If
        Call gauss.Add2Bazis(bazis(), row_indeces(), col_indeces()) 'Добавление полученного базиса в сводную матрицу
        Call add2BazisStrings
        If CreateSubTables = True Then 'Создание подтаблиц
            Set SubTableRange = SubTableRange.Offset(gauss.rows_count + 3, 0)
            Call gauss.createTables(SubTableRange, row_indeces(), col_indeces())
        End If
    Loop While comb.NextCombination = True 'Проверка наличия следующей комбинации индексов столбцов
End Sub

Public Sub Initialize() 'Начальная инициализация
    'Создание экземпляра clsGauss, задание параметров
    Set gauss = New clsGauss
    gauss.Initialize
    gauss.max_iterration = 2
    'Создание экземпляра clsCombinations, задание параметров
    Set comb = New clsCombinations
    comb.MaxElem = gauss.cols_count - 1
    comb.numElem = gauss.rows_count
    Call comb.init_comb
    Call init_index
End Sub

Public Sub bazisTable(ByRef table_range As Range) 'Создание таблицы базисов
    'Название таблицы
    With Range(Cells(table_range.Row - 2, table_range.Column - 1), Cells(table_range.Row - 2, table_range.Column + table_range.Columns.count - 2))
        .Merge
        .Value = "Все решения"
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Font.Bold = True
        .Font.Italic = True
        .Font.size = 12
    End With
    'Создание таблицы
    Call gauss.bazisTable(table_range, bazis(), bazis_titles())
End Sub


Public Sub getRefSolutions() 'Нахождение опорных решений в матрице базисов
    On Error GoTo ExitgetRefSolutions:
    Dim rows_count As Integer
    rows_count = UBound(bazis, 1)
    Dim Positive As Boolean
    For col = 1 To baz_Count
        Positive = True
        For Row = 1 To rows_count
            If bazis(Row, col) < 0 Then
                Positive = False
                Exit For
            End If
        Next
        If Positive = True Then
            Call add2RefSolutions(col, rows_count)
        End If
    Next
ExitgetRefSolutions:
Positive = Empty
rows_count = Empty
End Sub

Private Sub add2RefSolutions(ByRef col, ByRef rows_count As Integer)
    RefSolutionsCount = RefSolutionsCount + 1
    ReDim Preserve RefSolutions(rows_count, RefSolutionsCount)
    Call CopyMemory(RefSolutions(1, RefSolutionsCount), bazis(1, col), rows_count * 8)
    ReDim Preserve RefSolutionsTitles(1 To RefSolutionsCount)
    RefSolutionsTitles(RefSolutionsCount) = bazis_titles(col)
End Sub

Public Sub RefSolTable(ByRef table_range As Range) 'Создание таблицы опорных решений
    'Название таблицы
    With Range(Cells(table_range.Row - 2, table_range.Column - 1), Cells(table_range.Row - 2, table_range.Column + table_range.Columns.count - 2))
        .Merge
        .Value = "Опорные решения"
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Font.Bold = True
        .Font.Italic = True
        .Font.size = 12
    End With
    'Построение таблицы
    Call gauss.bazisTable(table_range, RefSolutions(), RefSolutionsTitles())
End Sub

Private Sub init_index() 'Начальная инициализация индексов матрицы
    ReDim col_indeces(1 To gauss.cols_count)
    col_indeces(gauss.cols_count) = gauss.cols_count
    Call gauss.init_rowIndeces(row_indeces())
End Sub

Private Sub add2BazisStrings() 'Создание заголовка базиса, добавление в массив заголовков
    baz_Count = baz_Count + 1
    ReDim Preserve bazis_titles(1 To baz_Count)
    For i = 1 To gauss.rows_count
        bazis_titles(baz_Count) = bazis_titles(baz_Count) & "a" & col_indeces(i) & ", "
    Next
    bazis_titles(baz_Count) = Mid(bazis_titles(baz_Count), 1, Len(bazis_titles(baz_Count)) - 2)
End Sub

Private Sub Class_Terminate()
    Set comb = Nothing
    Set gauss = Nothing
    CreateSubTables = Empty
    Set SubTableRange = Nothing
    Erase col_indeces()
    Erase bazis_titles()
    Erase row_indeces()
    Erase bazis()
    Erase RefSolutions()
    baz_Count = Empty
    RefSolutionsCount = Empty
End Sub
