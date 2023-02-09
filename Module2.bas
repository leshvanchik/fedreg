Attribute VB_Name = "Module2"
Sub Messages()
Dim m, k, head, deputy, all  As Integer
Dim default, loadDefault, gov As String
Dim cell_info, cell_gov, cell_message, cell_sphere As Range
Dim phrase As Variant
Dim dict_one, dict_two As Dictionary

Set dict_one = New Dictionary
dict_one.Add "*���������� ���������*", "F3"
dict_one.Add "*���������*", "F4"
dict_one.Add "*������������*", "F5"
dict_one.Add "*����� �� ��������*", "F6"
'dict_one.Add "*�������� ������*", "F8"
'dict_one.Add "*�����������*", "F9"
'dict_one.Add "*������*", "F10"
dict_one.Add "*������������� ����������� ���*", "F12"
dict_one.Add "*��������������� �������� ���*", "F13"
dict_one.Add "*���� ���*", "F14"
dict_one.Add "*�� ���*", "F15"
dict_one.Add "*����������� ��*", "F16"
dict_one.Add "*������ �����������*", "F17"
dict_one.Add "*������������ ��*", "F18"
dict_one.Add "*���������*", "F19"
dict_one.Add "*����*", "F20"
dict_one.Add "*���������*", "F22"
dict_one.Add "*�����������*", "F23"
dict_one.Add "*������*", "F24"
dict_one.Add "*���� (������, ����������� � �.�.)*", "F25"
dict_one.Add "*���������� ���, ������������ � ��������������� �� ����������� ������������ ���������*", "F26"
dict_one.Add "*���������� ��������*", "F27"
dict_one.Add "*����� �������� �� ������ ������������ � �������� �������*", "F93"
'dict_one.Add "*����������*", "F94"
dict_one.Add "*����������*", "F96"
dict_one.Add "*� ��� �����: ���� �������*", "F97"
dict_one.Add "*�� ����������*", "F98"
dict_one.Add "*��� ����� ������*", "F99"
dict_one.Add "*��������� ��� ������*", "F100"
dict_one.Add "*���������� �� �����������*", "F102"
dict_one.Add "*����������� � ������� �� �����*", "F103"
dict_one.Add "*����������� � ���������� �����*", "F104"
'dict_one.Add "*�� ������������*", "F105"

Set dict_two = New Dictionary
dict_two.Add "*�����������, ��������, ��������*", "F30"
dict_two.Add "*��������������� �����*", "F31"
dict_two.Add "*������ ���������������� ����������*", "F32"
dict_two.Add "*����������� �����*", "F33"
dict_two.Add "*������������� ���������. ������������� �����*", "F34"
dict_two.Add "*�������������� �������� ���� �� �������� ��������, �������� �����������, �����������, �����������, ���������� �������� � ���� ������*", "F35"
dict_two.Add "*���������� �����*", "F37"
dict_two.Add "*�����*", "F38"
dict_two.Add "*���� � ��������� ���������*", "F39"
dict_two.Add "*���������� ����������� � ���������� �����������*", "F40"
dict_two.Add "*�����������. �����. ��������*", "F42"
dict_two.Add "*����������� (�� ����������� �������������� ��������������)*", "F43"
dict_two.Add "*����� (�� ����������� �������������� �������������� � ������� �����)*", "F44"
dict_two.Add "*�������� (�� ����������� �������������� ��������������)*", "F45"
dict_two.Add "*�������� �������� ���������� (�� ����������� �������� ��������������)*", "F46"
dict_two.Add "*���������������. ���������� �������� � �����. ������*", "F48"
dict_two.Add "*��������������� (�� ����������� �������������� ��������������)*", "F49"
dict_two.Add "*���������� �������� � ����� (�� ����������� �������������� ��������������)*", "F50"
dict_two.Add "*������. ��������� (�� ����������� �������������� ��������������)*", "F51"
dict_two.Add "*���������*", "F53"
dict_two.Add "*�������*", "F54"
dict_two.Add "*������������� ������������*", "F56"
dict_two.Add "*��������������*", "F57"
dict_two.Add "*��������. �������� � �����������*", "F58"
dict_two.Add "*������������� ������� �������. ����������� ������������� ������� � ���������� (�� ����������� �������� ������������)*", "F59"
dict_two.Add "*�������������*", "F60"
dict_two.Add "*������������������ � �����������*", "F61"
dict_two.Add "*�������� ���������*", "F62"
dict_two.Add "*���������*", "F63"
dict_two.Add "*�����*", "F64"
dict_two.Add "*����������� ������������*", "F65"
dict_two.Add "*��������*", "F66"
dict_two.Add "*������������ �������*", "F67"
dict_two.Add "*������� ������������ ���������*", "F68"
dict_two.Add "*������������������� ������������. ���������� ����*", "F69"
dict_two.Add "*��������� ������� � ������ ���������� ��������� �����*", "F70"
dict_two.Add "*���������� � ��������������*", "F71"
dict_two.Add "*�������, ������������, ����������*", "F73"
dict_two.Add "*�������*", "F74"
dict_two.Add "*������������ � ������ ������������*", "F75"
dict_two.Add "*��������� �����. ���������� ���������*", "F76"
dict_two.Add "*����������*", "F77"
dict_two.Add "*�����������. ������ �������. ����������. ��������*", "F78"
dict_two.Add "*������*", "F80"
dict_two.Add "*����� ��������� ��������� ����������������*", "F81"
dict_two.Add "*�������� ����*", "F82"
dict_two.Add "*����������� ������� �������, ����������� �������� ������, ���������� �������� � �������� ����� (�� ����������� ����� ������������� �� ������)*", "F83"
dict_two.Add "*������������ ���������*", "F84"
dict_two.Add "*������ �������������, ���������� � ������� ����� (�������, �����������, ��������, ������)*", "F85"
dict_two.Add "*������� ���������. ���������������� ������ (� �������� �����)*", "F86"
dict_two.Add "*������� ��������� �� ����� � �������*", "F87"
dict_two.Add "*����������� ������������ (� �������� �����)*", "F88"
dict_two.Add "*������ ���������*", "F89"
dict_two.Add "*����������� ���������*", "F90"
dict_two.Add "*���������� �������� ������. ��������������� �� ��������� ��������� ����������������*", "F91"

'*****************************************************************************CODE*******************************************************************************************

Call UserForm1.Show

MsgBox "�������� ���� � ����������� � ����������� ������������ ��������� �������, ����������� � ������������ ����������� (�������� �� �����)", 48
default = Application.GetOpenFilename(Title:="�������� ����������� ����", FileFilter:="Excel Files(*.xls*), *xls*")

If default <> False Then
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    loadDefault = Dir(default)
    GetObject (loadDefault)
    
    With ThisWorkbook.Worksheets("�����������_����")
    gov = Workbooks(loadDefault).Worksheets("����1").Range("A2").Value
    For Each cell_info In ThisWorkbook.Worksheets("����������").Range("D1:D68")
        If cell_info Like gov & "*" Then
            For j = 3 To .Cells(Rows.Count, 1).End(xlUp).Row
                If .Range("E" & j).Offset(0, -1) <> "" Then
                    .Range("E" & j) = cell_info.Offset(0, -1).Value
                End If
            Next
            Exit For
        End If
    Next
    
    m = Workbooks(loadDefault).Worksheets(1).Cells(Rows.Count, 2).End(xlUp).Row
    For Each cell_message In Workbooks(loadDefault).Worksheets(1).Range("B1:B" & m)
        For i = 0 To dict_one.Count - 1
            If cell_message Like "*����������*" Or cell_message Like "*������ �� ������ �����*" Then
                .Range("F8") = .Range("F8").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like "*�����������*" Or cell_message Like "*����*" Then
                .Range("F9") = .Range("F9").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like "*������*" Or cell_message Like "*������ �����*" Then
                .Range("F10") = .Range("F10").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like "*����������*" Or cell_message Like "*�� ������������*" Then
                .Range("F94") = .Range("F94").Value + cell_message.Offset(0, 1).Value
                Exit For
            ElseIf cell_message Like dict_one.Keys()(i) Then
                .Range(dict_one.Items()(i)) = cell_message.Offset(0, 1).Value
                Exit For
            End If
        Next i
    Next cell_message
    
    k = Workbooks(loadDefault).Worksheets(2).Cells(Rows.Count, 2).End(xlUp).Row
    For Each cell_sphere In Workbooks(loadDefault).Worksheets(2).Range("B1:B" & k)
        For j = 0 To dict_two.Count - 1
            If cell_sphere Like dict_two.Keys()(j) Then
                .Range(dict_two.Items()(j)) = cell_sphere.Offset(0, 1).Value
                Exit For
            End If
        Next j
    Next cell_sphere
    
    If .Range("F26") = "" Then
        .Range("F26") = 0
    End If
         
    head = InputBox("���������� �������, �������� �� ������ ������� ������������� ��������������� ������", Title:="������� ����� �����")
    .Range("F107") = head
    deputy = InputBox("���������� �������, �������� �� ������ ������� ������������� ������������ ��������������� ������", Title:="������� ����� �����")
    .Range("F108") = deputy
    all = InputBox("���������� �������, �������� �� ������ ������� ������������� � ������������� ������������ ��������������� ������", Title:="������� ����� �����")
    .Range("F109") = all
    
    End With
    Workbooks(loadDefault).Close (False)
    
    Workbooks.Add
    ThisWorkbook.Worksheets("�����������_����").Range("A1:F109").Copy
    Workbooks("�����1").Worksheets(1).Range("A1:F109").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    ThisWorkbook.Worksheets("����1").Cells(1, 100).Clear
    ThisWorkbook.Worksheets("����1").Cells(1, 101).Clear
    ThisWorkbook.Close (False)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
End If

End Sub
