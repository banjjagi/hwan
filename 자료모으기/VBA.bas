Attribute VB_Name = "Module1"
Sub DrmFree()
'
' Macro1 Macro
'

    Dim �������� As String, ������� As String, ������ϸ� As String
    Dim �����ġ As String
    
    �������� = ActiveWorkbook.FullName
    �����ġ = ActiveWorkbook.Path
    ������ϸ� = "���" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    
    ActiveSheet.UsedRange.Select
    Selection.Copy
    Workbooks.Add
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Cells.Replace What:="[", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="]", Replacement:="", LookAt:=xlPart, SearchOrder:= _
        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

    ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ� & ".xlsx"
    
End Sub



Sub �������ڷ�ó��_�ű޿�()

    Dim �������� As String
    Dim ���� As Variant, �������� As Variant
    Dim �������� As String, ������� As String, ������ϸ� As String, �����ġ As String

'1. ���� ���� ��ȯ����

    �������� = "���� ���� (*.xlsx), *.xlsx"
    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="������ ���� ������ �����ϼ���", MultiSelect:=True)
    
    If IsArray(��������) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each ���� In ��������
        
            Workbooks.Open (����)

            �������� = ActiveWorkbook.FullName
            �����ġ = ActiveWorkbook.Path
            ������ϸ� = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Rows("1:5").Delete
         
            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=�����ġ & "\���\" & ������ϸ� & ".csv", FileFormat:=xlCSV
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close

        Next

    Else
        MsgBox "������ �������� �ʾҽ��ϴ�."
    End If

End Sub


Sub �������ڷ�ó��_�ű޿�()

    Dim �������� As String
    Dim ���� As Variant, �������� As Variant
    Dim �������� As String, ������� As String, ������ϸ� As String, �����ġ As String

'1. ���� ���� ��ȯ����

    �������� = "���� ���� (*.xlsx), *.xlsx"
    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="������ ���� ������ �����ϼ���", MultiSelect:=True)
    
    If IsArray(��������) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each ���� In ��������
        
            Workbooks.Open (����)

            �������� = ActiveWorkbook.FullName
            �����ġ = ActiveWorkbook.Path
            ������ϸ� = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Rows("1:4").Delete
            Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    
            Range("A1") = "����"
            Range("B1") = "����"
            Range("C1") = "�������"
            Range("D1") = "����"
            Range("E1") = "�Ϸù�ȣ"
            Range("F1") = "����ȣ"
            Range("G1") = "���ջ����"
            Range("H1") = "�������ֹι�ȣ"
            Range("I1") = "�������ֹ�SEQ"
            Range("J1") = "�����ڸ�"
            Range("K1") = "�������ֹι�ȣ"
            Range("L1") = "�������ֹ�SEQ"
            Range("M1") = "�����ڸ�"
            Range("N1") = "����ȯ�ޱݾ�"
        
            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=�����ġ & "\���\" & ������ϸ� & ".csv", FileFormat:=xlCSV
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close


        Next

    Else
        MsgBox "������ �������� �ʾҽ��ϴ�."
    End If

End Sub


Sub edi�ڷ�ó��_�ű޿�()

    Dim �������� As String
    Dim ���� As Variant, �������� As Variant
    Dim �������� As String, ������� As String, ������ϸ� As String, �����ġ As String
    Dim �ִ���� As Integer
    

'1. ���� ���� ��ȯ����




    �������� = "���� ���� (*.xlsx), *.xlsx"
    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="EDI ���� ������ �����ϼ���", MultiSelect:=True)
    
    If IsArray(��������) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each ���� In ��������
        
            Workbooks.Open (����)

            �������� = ActiveWorkbook.FullName
            �����ġ = ActiveWorkbook.Path
            ������ϸ� = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            
            If (Range("D4") = "��������") Then
                Columns("D").Insert Shift:=xlToRight
                Range("D4") = "������"
            End If
            
            �ִ���� = ActiveSheet.UsedRange.Rows.Count
            
            Columns("D").Insert Shift:=xlToRight
            Columns("D").NumberFormatLocal = "#"
            Range("D6").Select
            Range("D6") = "=mid($A$3,13,4)"
            Range("D6").Copy
            Range("D6").Offset(0, 0).Resize(�ִ���� - 5, 1).Select
            ActiveSheet.Paste
            Range("D6").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Range("D6").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
            Rows("5").Delete
            Rows("1:3").Delete
            
            
            

            With Columns("A:AL")
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            �ִ���� = ActiveSheet.UsedRange.Rows.Count
            
            Columns("AM:AR").Delete
            Columns("V").Delete
            Columns("A").Delete
            Columns("A:AI").NumberFormatLocal = "@"
            Columns("V").NumberFormatLocal = "#"
            Columns("E:E").Select
            Selection.Cut
            Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            
            Range("U1") = "���ޱݾ�"
            Range("T1") = "�ۼ�����"
            Range("P1") = "�������ֹι�ȣSEQ"
            Range("M1") = "�������ֹι�ȣSEQ"
            Range("AE1") = "����������˰��"
            Range("AF1") = "���¹�ȣ���˰��"
            Range("AG1") = "�������ֹι�ȣ���˰��"
            Range("AH1") = "�����ּ������˰��"
            Columns("A:AJ").EntireColumn.AutoFit
            
            '2. ��¥���İ���
            Columns("U:V").Insert Shift:=xlToRight
            Columns("U:V").NumberFormatLocal = "#"
            Range("U2").Select
            ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""YYYYMMDD"")"
            Range("U2").Copy
            Range("U2").Offset(0, 0).Resize(�ִ���� - 1, 2).Select
            ActiveSheet.Paste
            Range("U2:V2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Range("S2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Columns("U:V").Select
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
            
            Columns("Z:AA").Insert Shift:=xlToRight
            Columns("Z:AA").NumberFormatLocal = "#"
            Range("Z2").Select
            ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""YYYYMMDD"")"
            Range("Z2").Copy
            Range("Z2").Offset(0, 0).Resize(�ִ���� - 1, 2).Select
            ActiveSheet.Paste
            Range("Z2:AA2").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Range("X2").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Columns("Z:AA").Select
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
            
            Range("V1") = "����"
            Columns("U").Select
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
            
            Columns("F:F").Select
            Selection.Cut
            Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            Application.CutCopyMode = False
            
            Columns("C:F").Select
            Application.CutCopyMode = False
            Selection.Delete Shift:=xlToLeft
            
            
            
            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=�����ġ & "\���\" & ������ϸ� & ".csv", FileFormat:=xlCSV
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close


        Next

    Else
        MsgBox "������ �������� �ʾҽ��ϴ�."
    End If

End Sub



Sub �������ڷ�ó��_�ű޿�2()

    Dim �������� As String
    Dim ������� As String
    Dim ���� As Variant, �������� As Variant
    Dim ���ս�Ʈ As Worksheet
    Dim �۾����� As Workbook
    Dim ������ As Range
    Dim ������ġ As Range
    Dim �� As Long
    Dim ��1 As Long
    Dim ���ڵ庰��� As Long
    Dim �������� As String, ������� As String, ������ϸ� As String
    Dim �����ġ As String
    Dim �ִ���� As Long
    Dim �ִ뿭�� As Long
    
    

'1. �ؽ�Ʈ ���� ��ȯ����


    �������� = "�ؽ�Ʈ ���� (*.xlsx), *.xlsx"
    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾��� ���� ������ �����ϼ���", MultiSelect:=True)
    
    If IsArray(��������) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each ���� In ��������
        
            Workbooks.Open (����)
    
            '�����̸��� ���� �ؽ�Ʈ�� ����� �ٿ���

            �������� = ActiveWorkbook.FullName
            �����ġ = ActiveWorkbook.Path
            ������ϸ� = "���" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.NumberFormatLocal = "@"
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Selection.MergeCells = False

            �ִ���� = ActiveSheet.UsedRange.Rows.Count
            Rows(�ִ����).Delete
            Rows("1:8").Delete
            
            �ִ���� = ActiveSheet.UsedRange.Rows.Count
            
            For ��1 = 1 To �ִ����
                         
                    Cells(��1, 1).Resize(1, 14).Cut Destination:=Cells(5 * ((��1 - 1) \ 5) + 1, 14 * ((��1 - 1) Mod 5) + 1)
                    CutCopyMode = False
                    
            Next

            Range("A1").Select
            ActiveSheet.Sort.SortFields.Clear
            ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortTextAsNumbers
            With ActiveSheet.Sort
                .SetRange ActiveSheet.UsedRange
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With

            Columns("BR").Delete
            Columns("BM").Delete
            Columns("BH:BJ").Delete
            Columns("BC:BF").Delete
            Columns("BA").Delete
            Columns("AU").Delete
            Columns("AO:AR").Delete
            Columns("AM").Delete
            Columns("AK").Delete
            Columns("AF").Delete
            Columns("AA:AD").Delete
            Columns("X:Y").Delete
            Columns("S").Delete
            Columns("O:P").Delete
            Columns("M").Delete
            Columns("K").Delete
            Columns("I").Delete
            Columns("D:F").Delete
            Columns("B").Delete

            Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("A1") = "����"
            Range("B1") = "������ȣ"
            Range("C1") = "���ʰ����ݾ�"
            Range("D1") = "���������ݾ�"
            Range("E1") = "�����ڼ���"
            Range("F1") = "�������ֹι�ȣ"
            Range("G1") = "�������ֹι�ȣseq"
            Range("H1") = "��������ȣ"
            Range("I1") = "���ջ����"
            Range("J1") = "���������"
            Range("K1") = "�����뺸��"
            Range("L1") = "��ȿ�����"
            Range("M1") = "��������"
            Range("N1") = "���±���"
            Range("O1") = "�������ֹι�ȣ"
            Range("P1") = "�������ֹι�ȣseq"
            Range("Q1") = "�����ڸ�"
            Range("R1") = "���������"
            Range("S1") = "���������"
            Range("T1") = "��������ȭ��ȣ"
            Range("U1") = "���������ڵ�"
            Range("V1") = "��������ȣ"
            Range("W1") = "���ջ�����ڰ�"
            Range("X1") = "����������ڰ�"
            Range("Y1") = "�����ڱ���"
            Range("Z1") = "�ڰݱ���"
            Range("AA1") = "�ڰݻ���"
            Range("AB1") = "�������ڵ�����ȣ"
            Range("AC1") = "���������"
            Range("AD1") = "�����������ڰ�"
            Range("AE1") = "�������ȭ"
            Range("AF1") = "�����FAX"
            Range("AG1") = "����尡����"
            Range("AH1") = "�����Ż����"
            Range("AI1") = "EDI����"
            Range("AJ1") = "���������"
            
            Columns("A:AJ").NumberFormatLocal = "@"
            Columns("F:F").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.EntireRow.Delete

            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ� & ".xlsx"
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close

        Next

    Else
        MsgBox "������ �������� �ʾҽ��ϴ�."
    End If


End Sub



Sub ȯ��ó��_�ű޿�()

    Dim �������� As String
    Dim ������� As String
    Dim ���� As Variant, �������� As Variant
    Dim ���ս�Ʈ As Worksheet
    Dim �۾����� As Workbook
    Dim ������ As Range
    Dim ������ġ As Range
    Dim �� As Long
    Dim ��1 As Long
    Dim ���ڵ庰��� As Long
    Dim �������� As String, ������� As String, ������ϸ� As String
    Dim �����ġ As String
    Dim �ִ���� As Long
    Dim �ִ뿭�� As Long
    
    

'1. �ؽ�Ʈ ���� ��ȯ����


    �������� = "���� ���� (*.xls), *.xls"
    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾��� ���� ������ �����ϼ���", MultiSelect:=True)
    
    If IsArray(��������) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each ���� In ��������
        
            Workbooks.Open (����)
    
            '�����̸��� ���� �ؽ�Ʈ�� ����� �ٿ���

            �������� = ActiveWorkbook.FullName
            �����ġ = ActiveWorkbook.Path
            ������ϸ� = "���" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Selection.MergeCells = False
            
            Columns("A").Delete

            Rows(10).Delete
        
            �ִ���� = ActiveSheet.UsedRange.Rows.Count
            
            For ��1 = �ִ���� To 1 Step -1
                If ((��1 Mod 19) >= 1 And (��1 Mod 19) <= 9) Then
                Rows(��1).Delete
                End If
            Next
            
            
            
            For ��1 = 1 To �ִ����
                         
                    Cells(��1, 1).Resize(1, 14).Cut Destination:=Cells(2 * ((��1 - 1) \ 2) + 1, 14 * ((��1 - 1) Mod 2) + 1)
                    CutCopyMode = False
                    
            Next

            Range("A1").Select
            ActiveSheet.Sort.SortFields.Clear
            ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortTextAsNumbers
            With ActiveSheet.Sort
                .SetRange ActiveSheet.UsedRange
                .Header = xlNo
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With

            Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("A1") = "����"
            Range("B1") = "������ȣ"
            Range("C1") = "���ʰ����ݾ�"
            Range("D1") = "���������ݾ�"
            Range("E1") = "�����ڼ���"
            Range("F1") = "�������ֹι�ȣ"
            Range("G1") = "�������ֹι�ȣseq"
            Range("H1") = "��������ȣ"
            Range("I1") = "���ջ����"
            Range("J1") = "���������"
            Range("K1") = "�����뺸��"
            Range("L1") = "��ȿ�����"
            Range("M1") = "��������"
            Range("N1") = "���±���"
            Range("O1") = "�������ֹι�ȣ"
            Range("P1") = "�������ֹι�ȣseq"
            Range("Q1") = "�����ڸ�"
            Range("R1") = "���������"
            Range("S1") = "���������"
            Range("T1") = "��������ȭ��ȣ"
            Range("U1") = "���������ڵ�"
            Range("V1") = "��������ȣ"
            Range("W1") = "���ջ�����ڰ�"
            Range("X1") = "����������ڰ�"
            Range("Y1") = "�����ڱ���"
            Range("Z1") = "�ڰݱ���"
            Range("AA1") = "�ڰݻ���"
            Range("AB1") = "�������ڵ�����ȣ"
            Range("AC1") = "���������"
            Range("AD1") = "�����������ڰ�"
            Range("AE1") = "�������ȭ"
            Range("AF1") = "�����FAX"
            Range("AG1") = "����尡����"
            Range("AH1") = "�����Ż����"
            Range("AI1") = "EDI����"
            Range("AJ1") = "���������"
            
            Columns("C:C").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.EntireRow.Delete

            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ� & ".xlsx"
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close

        Next

    Else
        MsgBox "������ �������� �ʾҽ��ϴ�."
    End If


End Sub































































'Sub �������ڷ�ó��()
'
'    Dim �������� As String
'    Dim ������� As String
'    Dim ���� As Variant, �������� As Variant
'    Dim ���ս�Ʈ As Worksheet
'    Dim �۾����� As Workbook
'    Dim ������ As Range
'    Dim ������ġ As Range
'    Dim �� As Long
'    Dim ��1 As Long
'    Dim �������� As String, ������� As String, ������ϸ� As String
'    Dim �����ġ As String
'    Dim �ִ���� As Long
'    Dim �ִ뿭�� As Long
'
'
'
''1. �ؽ�Ʈ ���� ��ȯ����
'
'
'    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
'    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾��� �ؽ�Ʈ ������ �����ϼ���", MultiSelect:=True)
'
'        '    ������� = ThisWorkbook.Path & "\���޳���\"
'
'    If IsArray(��������) = True Then
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'        For Each ���� In ��������
'
'            Workbooks.OpenText Filename:=���� _
'            , Origin:=949, StartRow:=2, DataType:=xlDelimited, TextQualifier:= _
'            xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, _
'            Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'            Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'            Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2)), TrailingMinusNumbers:=True
'
'
''�����̸��� ���� �ؽ�Ʈ�� ����� �ٿ���
'
'            �������� = ActiveWorkbook.FullName
'            �����ġ = ActiveWorkbook.Path
'            ������ϸ� = "���" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
'
'            ActiveSheet.UsedRange.Select
'            Selection.Copy
'            Workbooks.Add
'
'            �ִ���� = ActiveSheet.UsedRange.Rows.Count
'            �ִ뿭�� = ActiveSheet.UsedRange.Columns.Count
'
'            For �� = 1 To �ִ����
'
'                Cells(��, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((�� - 1) \ 4) + 1, 11 * ((�� - 1) Mod 4) + 1)
'
'            Next
'
'            For ��1 = �ִ���� To 1 Step -1
'
'                If ��1 Mod 4 <> 1 Then
'                    Rows(��1).Delete
'                End If
'
'            Next
'
'            Application.ScreenUpdating = True
'            ActiveSheet.DisplayPageBreaks = True
'
'            ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ� & ".txt"
'
'        Next
'
'    Else
'        MsgBox "������ �������� �ʾҽ��ϴ�."
'    End If
'
''    If ���� <> "" Then
''
''        Application.ScreenUpdating = False
''
''            Set ���ս�Ʈ = ThisWorkbook.Worksheets(1)
''
''            ���ս�Ʈ.UsedRange.Offset(1).Delete Shift:=xlUp
''
''            Do
''
''                Set �۾����� = Workbooks.Open(Filename:=������� & ����)
''                Set ������ = �۾�����.Worksheets(1).UsedRange
''                With ������
''                    Set ������ = .Offset(1).Resize(.Rows.Count - 1)
''                End With
''
''                Set ������ġ = ���ս�Ʈ.Cells(Rows.Count, "A").End(xlUp).Offset(1)
''
''                ������.Copy ������ġ
''
''                �۾�����.Close SaveChanges:=False
''
''                ���� = Dir()
''
''            Loop While ���� <> ""
''
''        Application.ScreenUpdating = True
''
''    End If
''
'''2. �ؽ�Ʈ �������� ���� ����
''
''
''
''    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
''    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾� ���� ����")
''
''    If �������� <> False Then
''        Workbooks.OpenText Filename:=�������� _
''        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
''        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
''        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
''        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
''        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
''        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
''
''    Else
''        MsgBox "������ �������� �ʾҽ��ϴ�."
''    End If
''
''    Workbooks.Open (�����ġ & "\" & ������ϸ�)
''    Rows("1:11").Delete
''
''    �ִ���� = ActiveSheet.UsedRange.Rows.Count
''    �ִ뿭�� = ActiveSheet.UsedRange.Columns.Count
''
''    ActiveSheet.DisplayPageBreaks = False
''    Application.ScreenUpdating = False
''
''    For �� = 1 To �ִ����
''
''        Cells(��, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((�� - 1) \ 4) + 1, 11 * ((�� - 1) Mod 4) + 1)
''
''    Next
''
''    For ��1 = �ִ���� To 1 Step -1
''
''        If ��1 Mod 8 <> 1 Then
''            Rows(��1).Delete
''        End If
''
''    Next
''
''    '�� ����
''    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''    Cells.Replace What:="����", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''
''    Rows("1:1").Select
''    Selection.SpecialCells(xlCellTypeBlanks).Select
''    Selection.EntireColumn.Delete
''
''    Application.ScreenUpdating = True
''    ActiveSheet.DisplayPageBreaks = True
'
'
'
'End Sub
'
'
'
'
'
'Sub �����������ó��()
''
'' txt���Ϻҷ����� Macro
''
'
'    Dim �� As Long
'    Dim ��1 As Long
'    Dim �������� As String, ������� As String, ������ϸ� As String
'    Dim �����ġ As String
'    Dim �������� As String
'    Dim �������� As Variant
'    Dim �ִ���� As Long
'    Dim �ִ뿭�� As Long
'
'
'    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
'    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾� ���� ����")
'
'    If �������� <> False Then
'        Workbooks.OpenText Filename:=�������� _
'        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
'        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
'        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
'
'    Else
'        MsgBox "������ �������� �ʾҽ��ϴ�."
'    End If
'
'
''DRMǮ��
'
'    �������� = ActiveWorkbook.FullName
'    �����ġ = ActiveWorkbook.Path
'    ������ϸ� = "�����������(" & Date & ").xlsx"
'
'    ActiveSheet.UsedRange.Select
'    Selection.Copy
'    Workbooks.Add
'    Range("A1").Select
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=False
'    Cells.Replace What:="[", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="]", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ�
'
'  '�ϴ� ������ �� �� ���ϴ� ���� ���߿�
'
'    Workbooks.Open (�����ġ & "\" & ������ϸ�)
'    Rows("1:4").Delete
'
'    �ִ���� = ActiveSheet.UsedRange.Rows.Count
'    �ִ뿭�� = ActiveSheet.UsedRange.Columns.Count
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'    For �� = 1 To �ִ����
'
'        Cells(��, 1).Resize(1, 8).Cut Destination:=Cells(13 * ((�� - 1) \ 13) + 1, 8 * ((�� - 1) Mod 13) + 1)
'
'    Next
'
'    For ��1 = �ִ���� To 1 Step -1
'
'        If ��1 Mod 13 <> 1 Then
'            Rows(��1).Delete
'        End If
'
'    Next
'
'    '�� ����
'    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="����", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    Rows("1:1").Select
'
'    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Selection.EntireColumn.Delete
'
'    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'
'    Columns("X:AQ").Delete Shift:=xlToLeft
'    Columns("R:U").Delete Shift:=xlToLeft
'
'    Range("A1") = "ȯ�ް�����ȣ"
'    Range("B1") = "�������ֹι�ȣ"
'    Range("C1") = "�������ֹι�ȣseq"
'    Range("D1") = "�������ֹι�ȣ"
'    Range("E1") = "�������ֹι�ȣ"
'    Range("F1") = "�������ֹι�ȣseq"
'    Range("G1") = "�����ݾ�"
'    Range("H1") = "����"
'    Range("I1") = "�����"
'    Range("J1") = "�����ڼ���"
'    Range("K1") = "�����ּ���"
'    Range("L1") = "�����ڼ���"
'    Range("M1") = "����ȣ"
'    Range("N1") = "���¹�ȣ"
'    Range("O1") = "��������"
'    Range("P1") = "���忩��"
'    Range("Q1") = "����Ȯ�ο���"
'    Range("R1") = "���¸�"
'    Range("S1") = "��������"
'
''    Columns("A:A").Delete Shift:=xlToLeft
''    Columns("B:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
''    Columns("A:A").Select
''    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
''        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
''        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
''        :="-", FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, _
''        2)), TrailingMinusNumbers:=True
''
''    Range("A1") = "��������"
''    Range("B1") = "��������"
''    Range("C1") = "�������"
''    Range("D1") = "��������"
''    Range("E1") = "������ȣ"
''    Range("AK1") = "�ǿ����ּ���"
''
''    Application.ScreenUpdating = True
''    ActiveSheet.DisplayPageBreaks = True
''
''    Range("F2").Select
''    ActiveWindow.FreezePanes = True
''
''    Selection.AutoFilter
''    ActiveSheet.Range("$A$1:$AL$38").AutoFilter Field:=31, Criteria1:=Array( _
''        "���»���", "������", "������(�����ڵ�)", "="), Operator:=xlFilterValues
''
''    Columns("K:M").Select
''    Selection.EntireColumn.Hidden = True
''    Columns("O:P").Select
''    Selection.EntireColumn.Hidden = True
''    Columns("S:V").Select
''    Selection.EntireColumn.Hidden = True
''    Columns("AF:AJ").Select
''    Selection.EntireColumn.Hidden = True
''    Columns("AB:AB").Select
''    Selection.EntireColumn.Hidden = True
''    Columns("Z:AB").Select
''    Columns("Z:AB").EntireColumn.AutoFit
''    Columns("W:W").Select
''    Columns("W:W").EntireColumn.AutoFit
''
'
'End Sub
'
'
'
'Sub �����ް����۾�()
''
'' Macro3 Macro
'' �۾����� : 2014-10-29
'
''
'    Dim ���� As Range
'
'    ���� = ActiveSheet.UsedRange
'
'    ActiveWorkbook.Worksheets("�����ް�������").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("�����ް�������").Sort.SortFields.Add Key:=Range( _
'        "AB2:AB1001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'        xlSortTextAsNumbers
'    With ActiveWorkbook.Worksheets("�����ް�������").Sort
'        .SetRange ����
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'    ����.RemoveDuplicates Columns:=Array(6, 24, 25, 26 _
'        , 27), Header:=xlYes
'
'End Sub
'
'
'
'Sub �������¿�����ó��()
''
'' txt���Ϻҷ����� Macro
''
'
'    Dim �� As Long
'    Dim ��1 As Long
'    Dim �������� As String, ������� As String, ������ϸ� As String
'    Dim �����ġ As String
'    Dim �������� As String
'    Dim �������� As Variant
'    Dim �ִ���� As Long
'    Dim �ִ뿭�� As Long
'
'
'    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
'    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾� ���� ����")
'
'    If �������� <> False Then
'        Workbooks.OpenText Filename:=�������� _
'        , Origin:=949, StartRow:=9, DataType:=xlDelimited, TextQualifier:= _
'        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
'        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
'
'    Else
'        MsgBox "������ �������� �ʾҽ��ϴ�."
'    End If
'
'
''DRMǮ��
'
'    �������� = ActiveWorkbook.FullName
'    �����ġ = ActiveWorkbook.Path
'    ������ϸ� = "���" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
'
'    ActiveSheet.UsedRange.Select
'    Selection.Copy
'    Workbooks.Add
'    Range("A1").Select
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=False
'    Cells.Replace What:="[", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="]", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ� & ".xlsx"
'
'  '�ϴ� ������ �� �� ���ϴ� ���� ���߿�
'
'    Workbooks.Open (�����ġ & "\" & ������ϸ�)
'
'    �ִ���� = ActiveSheet.UsedRange.Rows.Count
'    �ִ뿭�� = ActiveSheet.UsedRange.Columns.Count
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'    For �� = 1 To �ִ����
'
'        Cells(��, 1).Resize(1, 12).Cut Destination:=Cells(3 * ((�� - 1) \ 3) + 1, 12 * ((�� - 1) Mod 3) + 1)
'
'    Next
'
'    For ��1 = �ִ���� To 1 Step -1
'
'        If ��1 Mod 3 <> 1 Then
'            Rows(��1).Delete
'        End If
'
'    Next
'
'    '�� ����
'    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="����", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    Rows("1:1").Select
'    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Selection.EntireColumn.Delete
'
'    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("A1") = "����"
'    Range("B1") = "�ڰݿ���"
'    Range("C1") = "�����ް�����ȣ"
'    Range("D1") = "�������ֹι�ȣ"
'    Range("E1") = "�����ڸ�"
'    Range("F1") = "�������ֹι�ȣ"
'    Range("G1") = "�����ڸ�"
'    Range("H1") = "�����ޱ�"
'    Range("I1") = "���޾ȳ�����"
'    Range("J1") = "�ȳ��뺸��"
'    Range("K1") = "���忩��"
'    Range("L1") = "���˰��"
'    Range("M1") = "�������ֹι�ȣ"
'    Range("N1") = "�����ڸ�"
'    Range("O1") = "�������ֹι�ȣ"
'    Range("P1") = "�����ָ�"
'    Range("Q1") = "�������"
'    Range("R1") = "���¹�ȣ"
'    Range("S1") = "Ÿ�ο���"
'    Range("T1") = "�ǿ����ָ�"
'    Range("U1") = "��û�ڸ�"
'    Range("V1") = "�������"
'
'    �ִ���� = ActiveSheet.UsedRange.Rows.Count
'    Rows(�ִ����).Delete
'
'    Range("A1").Select
'    Selection.AutoFilter
'    Range("a1").CurrentRegion.AutoFilter Field:=11, Criteria1:="N"
'
'    ActiveSheet.DisplayPageBreaks = True
'    Application.ScreenUpdating = True
'
'
'End Sub
'
'Sub �������ڷ���¿������()
'
'    Dim �������� As String
'    Dim ������� As String
'    Dim ���� As Variant, �������� As Variant
'    Dim ���ս�Ʈ As Worksheet
'    Dim �۾����� As Workbook
'    Dim ������ As Range
'    Dim ������ġ As Range
'    Dim �� As Long
'    Dim ��1 As Long
'    Dim ���ڵ庰��� As Long
'    Dim �������� As String, ������� As String, ������ϸ� As String
'    Dim �����ġ As String
'    Dim �ִ���� As Long
'    Dim �ִ뿭�� As Long
'
'
'
''1. �ؽ�Ʈ ���� ��ȯ����
'
'
'    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
'    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾��� �ؽ�Ʈ ������ �����ϼ���", MultiSelect:=True)
'
'        '    ������� = ThisWorkbook.Path & "\���޳���\"
'
'    If IsArray(��������) = True Then
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'
'        For Each ���� In ��������
'
'            Workbooks.OpenText Filename:=���� _
'            , Origin:=949, StartRow:=12, DataType:=xlDelimited, TextQualifier:= _
'            xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, _
'            Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'            Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 2), Array(8, 2), _
'            Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2)), TrailingMinusNumbers:=True
'
'
''�����̸��� ���� �ؽ�Ʈ�� ����� �ٿ���
'
'            �������� = ActiveWorkbook.FullName
'            �����ġ = ActiveWorkbook.Path
'            ������ϸ� = "���" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
'
'            ActiveSheet.UsedRange.Select
'            Selection.Copy
'
'            Workbooks.Add
'            Range("A1").Select
'            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
'            �ִ���� = ActiveSheet.UsedRange.Rows.Count
'            Rows(�ִ����).Delete
'
'            �ִ���� = ActiveSheet.UsedRange.Rows.Count
'
'            For ��1 = �ִ���� To 2 Step -1
'
'                ���ڵ庰��� = ��1 Mod 2
'                Select Case ���ڵ庰���
'
'                    Case 0
'
'                    Cells(��1, 1).Resize(1, 11).Cut Destination:=Cells(2 * ((��1 - 1) \ 2) + 1, 11 * ((��1 - 1) Mod 2) + 1)
'                    CutCopyMode = False
'
'
'                End Select
'
'            Next
'
'            Range("A1").Select
'            ActiveSheet.Sort.SortFields.Clear
'            ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), _
'                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'                xlSortTextAsNumbers
'            With ActiveSheet.Sort
'                .SetRange ActiveSheet.UsedRange
'                .Header = xlNo
'                .MatchCase = False
'                .Orientation = xlTopToBottom
'                .SortMethod = xlPinYin
'                .Apply
'            End With
'
'
'
'            Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'
'            Range("A1") = "����"
'            Range("B1") = "ȯ�ް�����ȣ"
'            Range("C1") = "�������ֹι�ȣ"
'            Range("D1") = "�������ֹι�ȣ�Ϸù�ȣ"
'            Range("E1") = "�������ֹι�ȣ"
'            Range("F1") = "�������ֹι�ȣ�Ϸù�ȣ"
'            Range("G1") = "������"
'            Range("H1") = "����ȣ"
'            Range("I1") = "��������"
'            Range("J1") = "�������"
'            Range("K1") = "�������ֹι�ȣ"
'            Range("L1") = "������"
'            Range("M1") = "��������"
'            Range("N1") = "�����ڼ���"
'            Range("O1") = "�����ڼ���"
'            Range("P1") = "���޻���"
'            Range("Q1") = "���ް����ݾ�"
'            Range("R1") = "���¹�ȣ"
'            Range("S1") = "�����ּ���"
'            Range("T1") = "�����ֿ� ����"
'
'            Application.ScreenUpdating = True
'            ActiveSheet.DisplayPageBreaks = True
'
'            ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ� & ".xlsx"
'            ActiveWorkbook.Close savechanges:=True
'
'            Application.Wait (Now + TimeValue("0:00:03"))
'            ActiveWorkbook.Close
'
'        Next
'
'    Else
'        MsgBox "������ �������� �ʾҽ��ϴ�."
'    End If
'
'
'
''    If ���� <> "" Then
''
''        Application.ScreenUpdating = False
''
''            Set ���ս�Ʈ = ThisWorkbook.Worksheets(1)
''
''            ���ս�Ʈ.UsedRange.Offset(1).Delete Shift:=xlUp
''
''            Do
''
''                Set �۾����� = Workbooks.Open(Filename:=������� & ����)
''                Set ������ = �۾�����.Worksheets(1).UsedRange
''                With ������
''                    Set ������ = .Offset(1).Resize(.Rows.Count - 1)
''                End With
''
''                Set ������ġ = ���ս�Ʈ.Cells(Rows.Count, "A").End(xlUp).Offset(1)
''
''                ������.Copy ������ġ
''
''                �۾�����.Close SaveChanges:=False
''
''                ���� = Dir()
''
''            Loop While ���� <> ""
''
''        Application.ScreenUpdating = True
''
''    End If
''
'''2. �ؽ�Ʈ �������� ���� ����
''
''
''
''    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
''    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾� ���� ����")
''
''    If �������� <> False Then
''        Workbooks.OpenText Filename:=�������� _
''        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
''        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
''        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
''        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
''        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
''        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
''
''    Else
''        MsgBox "������ �������� �ʾҽ��ϴ�."
''    End If
''
''    Workbooks.Open (�����ġ & "\" & ������ϸ�)
''    Rows("1:11").Delete
''
''    �ִ���� = ActiveSheet.UsedRange.Rows.Count
''    �ִ뿭�� = ActiveSheet.UsedRange.Columns.Count
''
''    ActiveSheet.DisplayPageBreaks = False
''    Application.ScreenUpdating = False
''
''    For �� = 1 To �ִ����
''
''        Cells(��, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((�� - 1) \ 4) + 1, 11 * ((�� - 1) Mod 4) + 1)
''
''    Next
''
''    For ��1 = �ִ���� To 1 Step -1
''
''        If ��1 Mod 8 <> 1 Then
''            Rows(��1).Delete
''        End If
''
''    Next
''
''    '�� ����
''    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''    Cells.Replace What:="����", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''
''    Rows("1:1").Select
''    Selection.SpecialCells(xlCellTypeBlanks).Select
''    Selection.EntireColumn.Delete
''
''    Application.ScreenUpdating = True
''    ActiveSheet.DisplayPageBreaks = True
'
'End Sub
'Sub edi���ó��()
''
'' txt���Ϻҷ����� Macro
''
'
'    Dim �� As Long
'    Dim ��1 As Long
'    Dim �������� As String, ������� As String, ������ϸ� As String
'    Dim �����ġ As String
'    Dim �������� As String
'    Dim �������� As Variant
'    Dim �ִ���� As Long
'    Dim �ִ뿭�� As Long
'
'
'    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
'    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾� ���� ����")
'
'    If �������� <> False Then
'        Workbooks.OpenText Filename:=�������� _
'        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
'        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
'        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
'
'    Else
'        MsgBox "������ �������� �ʾҽ��ϴ�."
'    End If
'
'
''DRMǮ��
'
'    �������� = ActiveWorkbook.FullName
'    �����ġ = ActiveWorkbook.Path
'    ������ϸ� = "���edi_" & Replace((Range("c5").Value & Range("h5").Value), ".", "") & ".xlsx"
'
'    ActiveSheet.UsedRange.Select
'    Selection.Copy
'    Workbooks.Add
'    Range("A1").Select
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=False
'    Cells.Replace What:="[", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="]", Replacement:="", LookAt:=xlPart, SearchOrder:= _
'        xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ�
'
'  '�ϴ� ������ �� �� ���ϴ� ���� ���߿�
'
'    Workbooks.Open (�����ġ & "\" & ������ϸ�)
'    Rows("1:11").Delete
'
'    �ִ���� = ActiveSheet.UsedRange.Rows.Count
'    �ִ뿭�� = ActiveSheet.UsedRange.Columns.Count
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'    For �� = 1 To �ִ����
'
'        Cells(��, 1).Resize(1, 11).Cut Destination:=Cells(8 * ((�� - 1) \ 8) + 1, 11 * ((�� - 1) Mod 8) + 1)
'
'    Next
'
'    For ��1 = �ִ���� To 1 Step -1
'
'        If ��1 Mod 8 <> 1 Then
'            Rows(��1).Delete
'        End If
'
'    Next
'
'    '�� ����
'    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="����", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    Rows("1:1").Select
'    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Selection.EntireColumn.Delete
'
'    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("A1") = "ȯ�ް�����ȣ"
'    Range("B1") = "����ȣ"
'    Range("C1") = "���ջ����"
'    Range("D1") = "������ȣ"
'    Range("E1") = "�����ڼ���"
'    Range("F1") = "�������ֹι�ȣ"
'    Range("G1") = "�������ֹι�ȣseq"
'    Range("H1") = "�ۼ�����"
'    Range("I1") = "��������"
'    Range("J1") = "�������ޱݾ�"
'    Range("K1") = "���������"
'    Range("L1") = "��������弼��"
'    Range("M1") = "�����ڼ���"
'    Range("N1") = "�������ֹι�ȣ"
'    Range("O1") = "�������ֹι�ȣseq"
'    Range("P1") = "��������"
'    Range("Q1") = "����"
'    Range("R1") = "��ȿ������"
'    Range("S1") = "�뺸���"
'    Range("T1") = "�����ڵ�"
'    Range("U1") = "���������"
'    Range("V1") = "�����ּ���"
'    Range("W1") = "�������ֹι�ȣ"
'    Range("X1") = "���¹�ȣ"
'    Range("Y1") = "�������ȭ��ȣ"
'    Range("Z1") = "�ȳ�����"
'    Range("AA1") = "��ûó�����"
'    Range("AB1") = "�ڷ����˻���"
'    Range("AC1") = "EDI����"
'    Range("AD1") = "�������Ȯ��"
'    Range("AE1") = "����������¹�ȣȮ��"
'    Range("AE1") = "���¹�ȣȮ��"
'    Range("AF1") = "�������ֹι�ȣȮ��"
'    Range("AG1") = "�����ּ���Ȯ��"
'    Range("AH1") = "�ǿ����ּ���"
'    Range("AI1") = "�����޽���"
'    Columns("AJ:AS").Delete Shift:=xlToLeft
'    Columns("Q:Q").Select
'    Selection.Cut
'    Columns("A:A").Select
'    Selection.Insert Shift:=xlToRight
'
'    Columns("A:A").Delete Shift:=xlToLeft
'    Columns("B:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'    Columns("A:A").Select
'    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
'        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
'        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
'        :="-", FieldInfo:=Array(Array(1, 2), Array(2, 2), Array(3, 2), Array(4, 2), Array(5, _
'        2)), TrailingMinusNumbers:=True
'
'    Range("A1") = "��������"
'    Range("B1") = "��������"
'    Range("C1") = "�������"
'    Range("D1") = "��������"
'    Range("E1") = "������ȣ"
'    Range("AK1") = "�ǿ����ּ���"
'
'    Application.ScreenUpdating = True
'    ActiveSheet.DisplayPageBreaks = True
'
'    Range("F2").Select
'    ActiveWindow.FreezePanes = True
'
'    Selection.AutoFilter
'    ActiveSheet.Range("$A$1:$AL$38").AutoFilter Field:=31, Criteria1:=Array( _
'        "���»���", "������", "������(�����ڵ�)", "�ڰ�������", "="), Operator:=xlFilterValues
'
'    Columns("K:M").Select
'    Selection.EntireColumn.Hidden = True
'    Columns("O:P").Select
'    Selection.EntireColumn.Hidden = True
'    Columns("S:V").Select
'    Selection.EntireColumn.Hidden = True
'    Columns("AF:AJ").Select
'    Selection.EntireColumn.Hidden = True
'    Columns("AB:AB").Select
'    Selection.EntireColumn.Hidden = True
'    Columns("Z:AB").Select
'    Columns("Z:AB").EntireColumn.AutoFit
'    Columns("W:W").Select
'    Columns("W:W").EntireColumn.AutoFit
'
''������ȣ ����
'
'    Range("H111").Select
'    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
'        ("H111"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortTextAsNumbers
'    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'End Sub
'
'Sub �������ڷ�ó��_����()
'
'    Dim �������� As String
'    Dim ������� As String
'    Dim ���� As Variant, �������� As Variant
'    Dim ���ս�Ʈ As Worksheet
'    Dim �۾����� As Workbook
'    Dim ������ As Range
'    Dim ������ġ As Range
'    Dim �� As Long
'    Dim ��1 As Long
'    Dim ���ڵ庰��� As Long
'    Dim �������� As String, ������� As String, ������ϸ� As String
'    Dim �����ġ As String
'    Dim �ִ���� As Long
'    Dim �ִ뿭�� As Long
'
'
'
''1. �ؽ�Ʈ ���� ��ȯ����
'
'
'    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
'    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾��� �ؽ�Ʈ ������ �����ϼ���", MultiSelect:=True)
'
'        '    ������� = ThisWorkbook.Path & "\���޳���\"
'
'    If IsArray(��������) = True Then
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'
'        For Each ���� In ��������
'
'            Workbooks.OpenText Filename:=���� _
'            , Origin:=949, StartRow:=2, DataType:=xlDelimited, TextQualifier:= _
'            xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, _
'            Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'            Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 2), Array(8, 2), _
'            Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2)), TrailingMinusNumbers:=True
'
'
''�����̸��� ���� �ؽ�Ʈ�� ����� �ٿ���
'
'            �������� = ActiveWorkbook.FullName
'            �����ġ = ActiveWorkbook.Path
'            ������ϸ� = "���" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
'
'            ActiveSheet.UsedRange.Select
'            Selection.Copy
'
'            Workbooks.Add
'            Range("A1").Select
'            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
'            �ִ���� = ActiveSheet.UsedRange.Rows.Count
'            Rows(�ִ����).Delete
'            Rows("1:6").Delete
'
'            �ִ���� = ActiveSheet.UsedRange.Rows.Count
'
'            For ��1 = �ִ���� To 2 Step -1
'
'                ���ڵ庰��� = ��1 Mod 4
'                Select Case ���ڵ庰���
'
'                    Case 2
'
'                    Cells(��1, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((��1 - 1) \ 4) + 1, 11 * ((��1 - 1) Mod 4) + 1)
'                    CutCopyMode = False
'
'                    Case 0
'                    Cells(��1, 1).Cut Destination:=Cells(4 * ((��1 - 1) \ 4) + 1, 23)
'                    CutCopyMode = False
'
'                End Select
'
'            Next
'
'            Range("A1").Select
'            ActiveSheet.Sort.SortFields.Clear
'            ActiveSheet.Sort.SortFields.Add Key:=Range("A1"), _
'                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'                xlSortTextAsNumbers
'            With ActiveSheet.Sort
'                .SetRange ActiveSheet.UsedRange
'                .Header = xlNo
'                .MatchCase = False
'                .Orientation = xlTopToBottom
'                .SortMethod = xlPinYin
'                .Apply
'            End With
'
'            Application.ScreenUpdating = True
'            ActiveSheet.DisplayPageBreaks = True
'
'            ActiveWorkbook.SaveAs Filename:=�����ġ & "\" & ������ϸ� & ".csv", FileFormat:=xlCSV
'            ActiveWorkbook.Close savechanges:=True
'
'            Application.Wait (Now + TimeValue("0:00:03"))
'            ActiveWorkbook.Close
'
'        Next
'
'    Else
'        MsgBox "������ �������� �ʾҽ��ϴ�."
'    End If
'
'
'
''    If ���� <> "" Then
''
''        Application.ScreenUpdating = False
''
''            Set ���ս�Ʈ = ThisWorkbook.Worksheets(1)
''
''            ���ս�Ʈ.UsedRange.Offset(1).Delete Shift:=xlUp
''
''            Do
''
''                Set �۾����� = Workbooks.Open(Filename:=������� & ����)
''                Set ������ = �۾�����.Worksheets(1).UsedRange
''                With ������
''                    Set ������ = .Offset(1).Resize(.Rows.Count - 1)
''                End With
''
''                Set ������ġ = ���ս�Ʈ.Cells(Rows.Count, "A").End(xlUp).Offset(1)
''
''                ������.Copy ������ġ
''
''                �۾�����.Close SaveChanges:=False
''
''                ���� = Dir()
''
''            Loop While ���� <> ""
''
''        Application.ScreenUpdating = True
''
''    End If
''
'''2. �ؽ�Ʈ �������� ���� ����
''
''
''
''    �������� = "�ؽ�Ʈ ���� (*.txt), *.txt"
''    �������� = Application.GetOpenFilename(filefilter:=��������, Title:="�۾� ���� ����")
''
''    If �������� <> False Then
''        Workbooks.OpenText Filename:=�������� _
''        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
''        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
''        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
''        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
''        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
''        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
''
''    Else
''        MsgBox "������ �������� �ʾҽ��ϴ�."
''    End If
''
''    Workbooks.Open (�����ġ & "\" & ������ϸ�)
''    Rows("1:11").Delete
''
''    �ִ���� = ActiveSheet.UsedRange.Rows.Count
''    �ִ뿭�� = ActiveSheet.UsedRange.Columns.Count
''
''    ActiveSheet.DisplayPageBreaks = False
''    Application.ScreenUpdating = False
''
''    For �� = 1 To �ִ����
''
''        Cells(��, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((�� - 1) \ 4) + 1, 11 * ((�� - 1) Mod 4) + 1)
''
''    Next
''
''    For ��1 = �ִ���� To 1 Step -1
''
''        If ��1 Mod 8 <> 1 Then
''            Rows(��1).Delete
''        End If
''
''    Next
''
''    '�� ����
''    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''    Cells.Replace What:="����", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''
''    Rows("1:1").Select
''    Selection.SpecialCells(xlCellTypeBlanks).Select
''    Selection.EntireColumn.Delete
''
''    Application.ScreenUpdating = True
''    ActiveSheet.DisplayPageBreaks = True
'
'End Sub
