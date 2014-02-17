Attribute VB_Name = "BazisSolve"
Private bazis As clsBazis '��������� ������ �����
Public isSubCalc As Boolean '��������� ��� ����������� ������������� ��������
Private bazisTableRange As Range '�������� ��������� �������

Public Sub BazisSolve() '��� �������� �������� �������
On Error GoTo ExitBazisSolve:
    If Adj.selectionCheck(False) = False Then
        GoTo ExitBazisSolve:
    End If
    Set bazisTableRange = Application.selection.Offset(1, Application.selection.Columns.count + 3)
    Set bazis = New clsBazis
    Call bazis.Initialize
    bazis.CreateSubTables = isSubCalc
    Call bazis.CalculateBazis
    Call bazis.bazisTable(bazisTableRange)
    Call bazis.getRefSolutions
    Set bazisTableRange = bazisTableRange.Offset(0, bazisTableRange.Columns.count + 2)
    Call bazis.RefSolTable(bazisTableRange)
ExitBazisSolve:
Set bazis = Nothing
Set bazisTableRange = Nothing
End Sub
