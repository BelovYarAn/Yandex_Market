Sub FindMinimumCost()
'������� ��������� ������
Call ClearList
'������� ������� �� ����������
Set wb = ThisWorkbook
With wb.Sheets(1)
    Set Cost = .Range("A1:Z1000").Find("���������")
    '���������� ��� ����������� ����� � ����������� ������
    Dim MinSum As Integer
    Dim MinRow As Integer
    '���������� ������ ������ ��� ����������� ��������
    MinSum = .Cells(2, Cost.Columns.Column).Value
    '� ����� ������� ���������� ��������
    For i = 2 To .Range(Cost.Columns.Column & ":" & Cost.Columns.Column).CurrentRegion.Rows.Count
        '���� �����, ���������� ���������� �������� � �������
        If .Cells(i, Cost.Columns.Column).Value < MinSum Then
            MinSum = .Cells(i, Cost.Columns.Column).Value
            MinRow = i
        End If
    Next i
    .Rows(MinRow).Select
End With
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .Color = 65535
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With
End Sub
Sub ClearList()
'������� ��������� ������
    ThisWorkbook.Sheets(1).Range("A2:Z1000").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
