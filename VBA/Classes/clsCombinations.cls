VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCombinations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Класс, генерирующий сочетания колонок. Сочетания из n по k элементов, в данном случае n = colsCount-1, k=rowsCount
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private comb() As Integer 'Массив, содержащий текущее сочетание
Private n As Integer 'Максимально возможный элемент сочетания
Private k As Integer 'Число элементов сочетания

Public Sub init_comb() 'Инициализация матрицы сочетания(генерится 1ое сочетание)
    ReDim comb(1 To k)
    For i = 1 To k
        comb(i) = i
    Next
End Sub

Public Function NextCombination() As Boolean 'Генерирует следующую комбинацию в лексикографическом порядке
    On Error GoTo ExitNextCombination:
    NextCombination = True
    Dim i As Integer
    i = k
    Do While i > 0
        If comb(i) < (n - (k - i)) Then
            Exit Do
        End If
        i = i - 1
    Loop
    If (i > 0) Then
        comb(i) = comb(i) + 1
        i = i + 1
        Do While i <= k
            comb(i) = comb(i - 1) + 1
            i = i + 1
        Loop
    Else
        NextCombination = False
    End If
'Освобождение памяти
ExitNextCombination:
i = Empty
End Function

Public Sub getColsIndex(ByRef cols_indeces() As Integer) 'Нумерует по новой колонки, добавляя к комбинации недостающие элементы
    On Error GoTo ExitgetColsIndex:
    Dim newElem As Integer
    newElem = 1
    Dim arrEnd As Integer
    arrEnd = k + 1
    'Заполняем массив элементами сочетания
    Call CopyMemory(cols_indeces(1), comb(1), k * 2)
    'Добиваем массив до начального элемента комбинации
    Do While newElem < comb(1)
        cols_indeces(arrEnd) = newElem
        newElem = newElem + 1
        arrEnd = arrEnd + 1
    Loop
    'Добиваем массив разницей между его элементами
    For i = 1 To k - 1
        newElem = comb(i) + 1
        Do While newElem < comb(i + 1)
            cols_indeces(arrEnd) = newElem
            newElem = newElem + 1
            arrEnd = arrEnd + 1
        Loop
    Next
    'Добиваем массив до максимального номера элемента
    newElem = newElem + 1
    Do While newElem <= n
        cols_indeces(arrEnd) = newElem
        newElem = newElem + 1
        arrEnd = arrEnd + 1
    Loop
'Освобождение памяти
ExitgetColsIndex:
newElem = Empty
arrEnd = Empty
End Sub

Public Property Let MaxElem(ByVal maxNum As Integer) 'Максимально возможный элемент сочетания(1..n)
    If maxNum > 0 Then
        n = maxNum
    End If
End Property

Public Property Let numElem(ByVal KElem As Integer)  'Число элементов сочетания
    If KElem > 0 Then
        k = KElem
    End If
End Property

Private Sub Class_Terminate()
    Erase comb()
    n = Empty
    k = Empty
End Sub
