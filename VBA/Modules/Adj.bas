Attribute VB_Name = "Adj"
'��������������� �����������

Public Function selectionCheck(ByVal square As Boolean) As Boolean ' �������� �� ������������ ��������� �������� �������
    selectionCheck = True
    If Application.selection.Cells.count <= 1 Then
        MsgBox ("������� �� ��������!")
        selectionCheck = False
        Exit Function
    End If
    If square = True Then
        If ((Application.selection.rows.count + 1) <> Application.selection.Columns.count) Then
            MsgBox ("������� �� ����������!!!!!")
            selectionCheck = False
            Exit Function
        End If
    End If
End Function

Public Sub Randomize() '��������� ��������� ���������� �������
    For Each Cell In selection
        Cell.Value = Int(200 * Rnd - 100)
    Next
End Sub

Public Sub RndAndCalcGauss() '������ ��� ������ - ��������� ������ � �����������
    Dim NowSelection As Range
    Set NowSelection = Application.selection
    Call Randomize
    Call GaussSolve.GaussSolve
    NowSelection.Select
'�������� ������
Set NowSelection = Nothing
End Sub

Public Sub SubCalculations() '������ ��� ��������, ����������� ������������� ��������
    BazisSolve.isSubCalc = Not BazisSolve.isSubCalc
End Sub

