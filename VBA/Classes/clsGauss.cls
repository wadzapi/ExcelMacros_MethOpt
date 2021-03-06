VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGauss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private gaussMatrix() As Double '������ ������� �������
Private orig_matrix() As Double '������ �������� �������
Private RowsCount As Integer '����� ����� �������� �������
Private colsCount As Integer '����� �������� �������� �������
Private resh As Double '�������� �������
Private disc() As Double '������ �������
Private totalDiscr As Double '��������� ����������� ����������
Private Const eps As Double = 0.000001 '�������� ���������� ��� ������ �������������
Private MaxIterr As Integer '������������ ����� ��������� ��������� ����������
Private bazis_col As Integer '����� �������� � ������� �������

Public Sub Initialize() '��������� �������������
    Call getSize
    Call init_Matr
End Sub

Private Sub init_Matr() '��������� ������������� ������� � ��������� ������
    ReDim orig_matrix(1 To RowsCount, 1 To colsCount)
    ReDim gaussMatrix(1 To RowsCount, 1 To colsCount)
    For Row = 1 To RowsCount
        For col = 1 To colsCount
            orig_matrix(Row, col) = Application.selection(Row, col).Value
        Next
    Next
End Sub

Private Sub fill_gaussMatr() '���������� ������� ������� ���������� ����������
    Call CopyMemory(gaussMatrix(1, 1), orig_matrix(1, 1), RowsCount * colsCount * 8)
End Sub

Private Sub getSize() '��������� ����������� �������
    RowsCount = Application.selection.rows.count
    colsCount = Application.selection.Columns.count
End Sub

Public Sub createTables(ByRef table As Range, ByRef rows_index() As Integer, ByRef cols_index() As Integer) '�������� ������
    Call New_Table(table, rows_index(), cols_index()) '�������� � ���������� ������� ����������� ��������� �������
    Call Discrepancy_table(table) '���������� ������� ������������
End Sub

Public Function CalculateMatrix(ByRef rows_index() As Integer, ByRef cols_index() As Integer) As Boolean '������� ������� �� ������ ������
        Call fill_gaussMatr
        If GaussMeth(rows_index(), cols_index()) = False Then
        CalculateMatrix = False
        Exit Function
    End If
        Call Discrepancy(rows_index(), cols_index()) '���������� ������� � ���������� �����������(����� ���� �������)
    If MaxIterr > 0 Then
        Call Iterration(rows_index(), cols_index()) '������������� ��������� ����������� ����������
    End If
CalculateMatrix = True
End Function

Private Function GaussMeth(ByRef rows_index() As Integer, ByRef cols_index() As Integer) As Boolean  '������� �� ������ ������
    '������ ������
    If GaussForward(rows_index(), cols_index()) = False Then
        GaussMeth = False
        Exit Function
    End If
    '�������� ������
    Call GaussBackward(rows_index(), cols_index())
GaussMeth = True
End Function
Private Sub New_Table(ByRef table As Range, ByRef rows_index() As Integer, ByRef cols_index() As Integer) '�������� ������� ��� ���������� �������
    On Error GoTo ExitNew_Table:
    '���������� ������� ���������� ���������� �������
    For Row = 1 To RowsCount
        For col = 1 To colsCount
            table(Row, col).Value = gaussMatrix(rows_index(Row), col)
        Next
    Next
    '�������������� �������
    With table
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .BorderAround ColorIndex:=5, Weight:=xlThick
    End With
    '�������� ����������
    Dim headerRange As Range '�������,���������� ��������� �������
    Set headerRange = table.Resize(1, table.Columns.count).Offset(-1, 0)
    For col = 1 To colsCount - 1
        headerRange(1, col).Value = "a" & col
    Next
    headerRange(1, colsCount).Value = "b"
    '�������������� ���������� �������
    With headerRange
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Font.Bold = True
        .Font.Italic = True
        .Font.size = 12
        .RowHeight = 19.5
        .BorderAround ColorIndex:=5, Weight:=xlThick
        .Interior.ColorIndex = 4
    End With
'������������ ������
ExitNew_Table:
Set headerRange = Nothing
End Sub

Private Function GaussForward(ByRef rows_index() As Integer, ByRef cols_index() As Integer) As Boolean '������ ��� ������ ������
    On Error GoTo ExitGaussForward:
    For shag = 1 To RowsCount
        '����� ��������� ��������
        If findR(shag, rows_index(), cols_index()) = False Then
            GaussForward = False
            GoTo ExitGaussForward:
        End If
        resh = gaussMatrix(rows_index(shag), cols_index(shag))
        For i = 1 To colsCount
            gaussMatrix(rows_index(shag), cols_index(i)) = gaussMatrix(rows_index(shag), cols_index(i)) / resh
        Next
        For Row = (shag + 1) To RowsCount
            resh = gaussMatrix(rows_index(Row), cols_index(shag))
            For col = 1 To colsCount
                gaussMatrix(rows_index(Row), cols_index(col)) = gaussMatrix(rows_index(Row), cols_index(col)) - resh * gaussMatrix(rows_index(shag), cols_index(col))
            Next
        Next
    Next
GaussForward = True
'������������ ������ � �����
ExitGaussForward:
resh = Empty
End Function

Private Sub GaussBackward(ByRef rows_index() As Integer, ByRef cols_index() As Integer) '�������� ��� ������ ������
    On Error GoTo ExitGaussBackward:
    For shag = RowsCount To 1 Step -1
        For Row = (shag - 1) To 1 Step -1
        resh = gaussMatrix(rows_index(Row), cols_index(shag))
            For col = (colsCount) To 1 Step -1
                gaussMatrix(rows_index(Row), cols_index(col)) = gaussMatrix(rows_index(Row), cols_index(col)) - resh * gaussMatrix(rows_index(shag), cols_index(col))
            Next
        Next
    Next
'������������ ������
ExitGaussBackward:
resh = Empty
End Sub

Private Function findR(ByVal shag As Integer, ByRef rows_index() As Integer, ByRef cols_index() As Integer) As Boolean ' ����� ��������� ��-�� �� �������
    On Error GoTo ExitfindR:
    '���� ������ � ������������ �� ������ �������������
    Dim max_abs As Double
    Dim cur_abs As Double
    Dim max_abs_row As Integer
    For Row = shag To RowsCount
        cur_abs = Abs(gaussMatrix(rows_index(Row), cols_index(shag)))
        If cur_abs >= max_abs Then
            max_abs = cur_abs
            max_abs_row = Row
        End If
    Next
    '�������� ������� �� ������������ � �� ������������ �������
    If (max_abs < eps) Then
        If Abs(gaussMatrix(max_abs_row, cols_index(colsCount))) > eps Then
            MsgBox ("������� ��������� �����������")
            findR = False
            GoTo ExitfindR:
        Else
            MsgBox ("������� ��������� ����� ��������� �������")
            findR = False
            GoTo ExitfindR:
        End If
    End If
    '������������ ������, � ������������ � ����������� ������������
    If (rows_index(shag) <> max_abs_row) Then
        Dim buff As Double
        buff = rows_index(shag)
        rows_index(shag) = rows_index(max_abs_row)
        rows_index(max_abs_row) = buff
    End If
findR = True
'������������ ������ � �����
ExitfindR:
max_abs = Empty
cur_abs = Empty
max_abs_row = Empty
buff = Empty
End Function

Private Sub Discrepancy(ByRef rows_index() As Integer, ByRef cols_index() As Integer) '���������� �������
'������� ���������� ���������� ������������
Erase disc
totalDiscr = Empty
ReDim disc(1 To RowsCount)
    '�������� �������� � ��������
    For Row = 1 To RowsCount
        For col = 1 To RowsCount
            disc(rows_index(Row)) = disc(rows_index(Row)) + orig_matrix(rows_index(Row), cols_index(col)) * gaussMatrix(rows_index(col), cols_index(colsCount))
        Next
    Next
    '��������� ������� � ��������� �����������
    For Row = 1 To RowsCount
        disc(rows_index(Row)) = orig_matrix(rows_index(Row), colsCount) - disc(rows_index(Row))
        totalDiscr = totalDiscr + (disc(rows_index(Row))) ^ 2
    Next
End Sub

Private Sub Discrepancy_table(ByRef table As Range) '�������� � ���������� ������� �������
    On Error GoTo ExitDiscrepancy_table:
    Dim discr_range As Range
    Set discr_range = table.Offset(0, colsCount + 1).Resize(RowsCount, 1)
    '��������� - ����������, ��������������
    With Cells(discr_range.Row - 1, discr_range.Column)
        .Value = "�������"
        .HorizontalAlignment = xlHAlignCenter
        .Font.Bold = True
        .Font.size = 12
    End With
    With Cells(discr_range.Row + RowsCount, discr_range.Column - 1)
        .Value = "R(x)"
        .HorizontalAlignment = xlHAlignRight
        .Font.size = 12
        .Font.Bold = True
    End With
    '���������� �������
    For Row = 1 To RowsCount
        discr_range(Row, 1).Value = disc(Row)
    Next
    '���������� ��������� �����������, ��������������
    With Cells(discr_range.Row + RowsCount, discr_range.Column)
        .Value = totalDiscr
        .Interior.ColorIndex = 3
        .BorderAround ColorIndex:=5, Weight:=xlThick
    End With
    '�������������� ������� �������
    With discr_range
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .ColumnWidth = 13
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .BorderAround ColorIndex:=5, Weight:=xlThick
    End With
'�������� ������
ExitDiscrepancy_table:
Set discr_range = Nothing
End Sub

Public Sub bazisTable(ByRef table_range As Range, ByRef values() As Double, ByRef bazis_titles() As String) '�������� ������� ������� �������
    On Error GoTo ExitBazisTable:
    Dim rows_count As Integer
    rows_count = UBound(values(), 2)
    Dim cols_count As Integer
    cols_count = UBound(values(), 1)
    '���������� ������� ������� ����������
    Set table_range = table_range.Resize(rows_count, cols_count)
    For Row = 1 To cols_count
        For col = 1 To rows_count
            table_range(col, Row).Value = values(Row, col)
        Next
    Next
    '���������� �������� �������
    Dim titles As Range
    Set titles = table_range.Resize(rows_count, 1).Offset(0, -1)
    For Row = 1 To rows_count
        titles(Row, 1).Value = bazis_titles(Row)
    Next
    '���������� ����� �������
    Dim header As Range
    Set header = table_range.Resize(1, cols_count + 1).Offset(-1, -1)
    header(1, 1) = "����� / ��������"
    For col = 1 To cols_count
        header(1, col + 1) = "x" & col
    Next
    '��������������
    '�������������� ������� ��������
    With table_range
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .ColumnWidth = 13
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .BorderAround ColorIndex:=10, Weight:=xlThick
    End With
    '�������������� ��������� �������
    With titles
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Font.Bold = True
        .Font.Italic = True
        .Font.size = 12
        .ColumnWidth = 23
        .BorderAround ColorIndex:=10, Weight:=xlThick
        .Interior.ColorIndex = 6
    End With
    '�������������� ����� �������
    With header
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
        .Font.Bold = True
        .Font.Italic = True
        .Font.size = 12
        .BorderAround ColorIndex:=10, Weight:=xlThick
        .Interior.ColorIndex = 7
    End With
ExitBazisTable:
'������������ ������
rows_count = Empty
cols_count = Empty
Set titles = Nothing
Set header = Nothing
End Sub


Private Sub Iterration(ByRef rows_index() As Integer, ByRef cols_index() As Integer) '������������ ��������� ����������� ����������
    On Error GoTo ExitIterration:
    Dim summAi As Double
    '���������� maxIterr �����������
    For Iterrat = 1 To MaxIterr
        For Row = 1 To RowsCount
            For col = 1 To RowsCount
            summAi = summAi + orig_matrix(Row, col)
            Next
            If summAi <> 0 Then
                gaussMatrix(rows_index(Row), cols_index(colsCount)) = gaussMatrix(rows_index(Row), cols_index(colsCount)) + disc(Row) / summAi
            End If
        Next
        '���������� ����������� �������(�����������)
        Call Discrepancy(rows_index(), cols_index())
    Next
'������������ ������
ExitIterration:
summAi = Empty
End Sub
Public Sub init_rowIndeces(ByRef index() As Integer) '��������� ������������� �������� ����� �������

    ReDim index(1 To RowsCount)
    For i = 1 To RowsCount
        index(i) = i
    Next
    
End Sub

Public Sub init_colIndeces(ByRef index() As Integer) '��������� ������������� �������� �������� �������

    ReDim index(1 To colsCount)
    For i = 1 To colsCount
        index(i) = i
    Next
    
End Sub

Public Property Get rows_count() As Integer '��� ��������� ����� ����� �������
    rows_count = RowsCount
End Property

Public Property Get cols_count() As Integer '��� ��������� ����� �������� �������
    cols_count = colsCount
End Property

Public Property Let max_iterration(ByVal max_iterr As Integer) '��� ������� ����� ��������� ��������� ����������
    MaxIterr = max_iterr
End Property

Public Sub Add2Bazis(ByRef bazis() As Double, ByRef rows_index() As Integer, ByRef cols_index() As Integer) '���������� ����������� � ������� ������� �������
    bazis_col = bazis_col + 1
    ReDim Preserve bazis(1 To colsCount - 1, 1 To bazis_col)
    For i = 1 To RowsCount
        bazis(cols_index(i), bazis_col) = gaussMatrix(rows_index(i), colsCount)
    Next
End Sub

Private Sub Class_Terminate() '�������� �� ����������� ���������� ������
 MaxIterr = Empty
 bazis_col = Empty
 totalDiscr = Empty
 colsCount = Empty
 RowsCount = Empty
 Erase gaussMatrix()
 Erase orig_matrix()
 Erase disc()
End Sub
