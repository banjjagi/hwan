Attribute VB_Name = "Module1"
Sub DrmFree()
'
' Macro1 Macro
'

    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
    Dim 백업위치 As String
    
    현재파일 = ActiveWorkbook.FullName
    백업위치 = ActiveWorkbook.Path
    백업파일명 = "백업" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
    
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

    ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명 & ".xlsx"
    
End Sub



Sub 기지급자료처리_신급여()

    Dim 파일형식 As String
    Dim 파일 As Variant, 선택파일 As Variant
    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String, 백업위치 As String

'1. 엑셀 파일 순환가능

    파일형식 = "엑셀 파일 (*.xlsx), *.xlsx"
    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="기지급 엑셀 파일을 선택하세요", MultiSelect:=True)
    
    If IsArray(선택파일) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each 파일 In 선택파일
        
            Workbooks.Open (파일)

            현재파일 = ActiveWorkbook.FullName
            백업위치 = ActiveWorkbook.Path
            백업파일명 = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Rows("1:5").Delete
         
            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=백업위치 & "\백업\" & 백업파일명 & ".csv", FileFormat:=xlCSV
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close

        Next

    Else
        MsgBox "파일을 선택하지 않았습니다."
    End If

End Sub


Sub 미지급자료처리_신급여()

    Dim 파일형식 As String
    Dim 파일 As Variant, 선택파일 As Variant
    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String, 백업위치 As String

'1. 엑셀 파일 순환가능

    파일형식 = "엑셀 파일 (*.xlsx), *.xlsx"
    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="미지급 엑셀 파일을 선택하세요", MultiSelect:=True)
    
    If IsArray(선택파일) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each 파일 In 선택파일
        
            Workbooks.Open (파일)

            현재파일 = ActiveWorkbook.FullName
            백업위치 = ActiveWorkbook.Path
            백업파일명 = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            Rows("1:4").Delete
            Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    
            Range("A1") = "직역"
            Range("B1") = "지사"
            Range("C1") = "결정년월"
            Range("D1") = "종별"
            Range("E1") = "일련번호"
            Range("F1") = "증번호"
            Range("G1") = "통합사업장"
            Range("H1") = "수진자주민번호"
            Range("I1") = "수진자주민SEQ"
            Range("J1") = "수진자명"
            Range("K1") = "가입자주민번호"
            Range("L1") = "가입자주민SEQ"
            Range("M1") = "가입자명"
            Range("N1") = "최종환급금액"
        
            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=백업위치 & "\백업\" & 백업파일명 & ".csv", FileFormat:=xlCSV
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close


        Next

    Else
        MsgBox "파일을 선택하지 않았습니다."
    End If

End Sub


Sub edi자료처리_신급여()

    Dim 파일형식 As String
    Dim 파일 As Variant, 선택파일 As Variant
    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String, 백업위치 As String
    Dim 최대행수 As Integer
    

'1. 엑셀 파일 순환가능




    파일형식 = "엑셀 파일 (*.xlsx), *.xlsx"
    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="EDI 엑셀 파일을 선택하세요", MultiSelect:=True)
    
    If IsArray(선택파일) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each 파일 In 선택파일
        
            Workbooks.Open (파일)

            현재파일 = ActiveWorkbook.FullName
            백업위치 = ActiveWorkbook.Path
            백업파일명 = Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            
            
            If (Range("D4") = "접수구분") Then
                Columns("D").Insert Shift:=xlToRight
                Range("D4") = "저장결과"
            End If
            
            최대행수 = ActiveSheet.UsedRange.Rows.Count
            
            Columns("D").Insert Shift:=xlToRight
            Columns("D").NumberFormatLocal = "#"
            Range("D6").Select
            Range("D6") = "=mid($A$3,13,4)"
            Range("D6").Copy
            Range("D6").Offset(0, 0).Resize(최대행수 - 5, 1).Select
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
            
            최대행수 = ActiveSheet.UsedRange.Rows.Count
            
            Columns("AM:AR").Delete
            Columns("V").Delete
            Columns("A").Delete
            Columns("A:AI").NumberFormatLocal = "@"
            Columns("V").NumberFormatLocal = "#"
            Columns("E:E").Select
            Selection.Cut
            Columns("A:A").Select
            Selection.Insert Shift:=xlToRight
            
            Range("U1") = "지급금액"
            Range("T1") = "작성일자"
            Range("P1") = "가입자주민번호SEQ"
            Range("M1") = "수진자주민번호SEQ"
            Range("AE1") = "금융기관점검결과"
            Range("AF1") = "계좌번호점검결과"
            Range("AG1") = "예금주주민번호점검결과"
            Range("AH1") = "예금주성명점검결과"
            Columns("A:AJ").EntireColumn.AutoFit
            
            '2. 날짜서식관련
            Columns("U:V").Insert Shift:=xlToRight
            Columns("U:V").NumberFormatLocal = "#"
            Range("U2").Select
            ActiveCell.FormulaR1C1 = "=TEXT(RC[-2],""YYYYMMDD"")"
            Range("U2").Copy
            Range("U2").Offset(0, 0).Resize(최대행수 - 1, 2).Select
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
            Range("Z2").Offset(0, 0).Resize(최대행수 - 1, 2).Select
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
            
            Range("V1") = "차수"
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
            
            ActiveWorkbook.SaveAs Filename:=백업위치 & "\백업\" & 백업파일명 & ".csv", FileFormat:=xlCSV
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close


        Next

    Else
        MsgBox "파일을 선택하지 않았습니다."
    End If

End Sub



Sub 미지급자료처리_신급여2()

    Dim 파일형식 As String
    Dim 대상폴더 As String
    Dim 파일 As Variant, 선택파일 As Variant
    Dim 통합시트 As Worksheet
    Dim 작업파일 As Workbook
    Dim 대상범위 As Range
    Dim 복사위치 As Range
    Dim 행 As Long
    Dim 행1 As Long
    Dim 레코드별행수 As Long
    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
    Dim 백업위치 As String
    Dim 최대행수 As Long
    Dim 최대열수 As Long
    
    

'1. 텍스트 파일 순환가능


    파일형식 = "텍스트 파일 (*.xlsx), *.xlsx"
    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업할 엑셀 파일을 선택하세요", MultiSelect:=True)
    
    If IsArray(선택파일) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each 파일 In 선택파일
        
            Workbooks.Open (파일)
    
            '파일이름은 예전 텍스트에 백업만 붙여서

            현재파일 = ActiveWorkbook.FullName
            백업위치 = ActiveWorkbook.Path
            백업파일명 = "백업" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.NumberFormatLocal = "@"
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Selection.MergeCells = False

            최대행수 = ActiveSheet.UsedRange.Rows.Count
            Rows(최대행수).Delete
            Rows("1:8").Delete
            
            최대행수 = ActiveSheet.UsedRange.Rows.Count
            
            For 행1 = 1 To 최대행수
                         
                    Cells(행1, 1).Resize(1, 14).Cut Destination:=Cells(5 * ((행1 - 1) \ 5) + 1, 14 * ((행1 - 1) Mod 5) + 1)
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
            Range("A1") = "순번"
            Range("B1") = "결정번호"
            Range("C1") = "최초결정금액"
            Range("D1") = "최종결정금액"
            Range("E1") = "수진자성명"
            Range("F1") = "수진자주민번호"
            Range("G1") = "수진자주민번호seq"
            Range("H1") = "결정증번호"
            Range("I1") = "통합사업장"
            Range("J1") = "결정사업장"
            Range("K1") = "최종통보일"
            Range("L1") = "시효기산일"
            Range("M1") = "변동사유"
            Range("N1") = "상태구분"
            Range("O1") = "가입자주민번호"
            Range("P1") = "가입자주민번호seq"
            Range("Q1") = "가입자명"
            Range("R1") = "최종취득일"
            Range("S1") = "최종상실일"
            Range("T1") = "가입자전화번호"
            Range("U1") = "최종지사코드"
            Range("V1") = "최종증번호"
            Range("W1") = "통합사업장자격"
            Range("X1") = "결정사업장자격"
            Range("Y1") = "가입자구분"
            Range("Z1") = "자격구분"
            Range("AA1") = "자격상태"
            Range("AB1") = "가입자핸드폰번호"
            Range("AC1") = "최종지사명"
            Range("AD1") = "최종사업장명자격"
            Range("AE1") = "사업장전화"
            Range("AF1") = "사업장FAX"
            Range("AG1") = "사업장가입일"
            Range("AH1") = "사업장탈퇴일"
            Range("AI1") = "EDI적용"
            Range("AJ1") = "사업장적용"
            
            Columns("A:AJ").NumberFormatLocal = "@"
            Columns("F:F").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.EntireRow.Delete

            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명 & ".xlsx"
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close

        Next

    Else
        MsgBox "파일을 선택하지 않았습니다."
    End If


End Sub



Sub 환입처리_신급여()

    Dim 파일형식 As String
    Dim 대상폴더 As String
    Dim 파일 As Variant, 선택파일 As Variant
    Dim 통합시트 As Worksheet
    Dim 작업파일 As Workbook
    Dim 대상범위 As Range
    Dim 복사위치 As Range
    Dim 행 As Long
    Dim 행1 As Long
    Dim 레코드별행수 As Long
    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
    Dim 백업위치 As String
    Dim 최대행수 As Long
    Dim 최대열수 As Long
    
    

'1. 텍스트 파일 순환가능


    파일형식 = "엑셀 파일 (*.xls), *.xls"
    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업할 엑셀 파일을 선택하세요", MultiSelect:=True)
    
    If IsArray(선택파일) = True Then
    
    ActiveSheet.DisplayPageBreaks = False
    Application.ScreenUpdating = False

           
        For Each 파일 In 선택파일
        
            Workbooks.Open (파일)
    
            '파일이름은 예전 텍스트에 백업만 붙여서

            현재파일 = ActiveWorkbook.FullName
            백업위치 = ActiveWorkbook.Path
            백업파일명 = "백업" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
            
            ActiveSheet.UsedRange.Select
            Selection.Copy
            
            Workbooks.Add
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Selection.MergeCells = False
            
            Columns("A").Delete

            Rows(10).Delete
        
            최대행수 = ActiveSheet.UsedRange.Rows.Count
            
            For 행1 = 최대행수 To 1 Step -1
                If ((행1 Mod 19) >= 1 And (행1 Mod 19) <= 9) Then
                Rows(행1).Delete
                End If
            Next
            
            
            
            For 행1 = 1 To 최대행수
                         
                    Cells(행1, 1).Resize(1, 14).Cut Destination:=Cells(2 * ((행1 - 1) \ 2) + 1, 14 * ((행1 - 1) Mod 2) + 1)
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
            Range("A1") = "순번"
            Range("B1") = "결정번호"
            Range("C1") = "최초결정금액"
            Range("D1") = "최종결정금액"
            Range("E1") = "수진자성명"
            Range("F1") = "수진자주민번호"
            Range("G1") = "수진자주민번호seq"
            Range("H1") = "결정증번호"
            Range("I1") = "통합사업장"
            Range("J1") = "결정사업장"
            Range("K1") = "최종통보일"
            Range("L1") = "시효기산일"
            Range("M1") = "변동사유"
            Range("N1") = "상태구분"
            Range("O1") = "가입자주민번호"
            Range("P1") = "가입자주민번호seq"
            Range("Q1") = "가입자명"
            Range("R1") = "최종취득일"
            Range("S1") = "최종상실일"
            Range("T1") = "가입자전화번호"
            Range("U1") = "최종지사코드"
            Range("V1") = "최종증번호"
            Range("W1") = "통합사업장자격"
            Range("X1") = "결정사업장자격"
            Range("Y1") = "가입자구분"
            Range("Z1") = "자격구분"
            Range("AA1") = "자격상태"
            Range("AB1") = "가입자핸드폰번호"
            Range("AC1") = "최종지사명"
            Range("AD1") = "최종사업장명자격"
            Range("AE1") = "사업장전화"
            Range("AF1") = "사업장FAX"
            Range("AG1") = "사업장가입일"
            Range("AH1") = "사업장탈퇴일"
            Range("AI1") = "EDI적용"
            Range("AJ1") = "사업장적용"
            
            Columns("C:C").Select
            Selection.SpecialCells(xlCellTypeBlanks).Select
            Selection.EntireRow.Delete

            Application.ScreenUpdating = True
            ActiveSheet.DisplayPageBreaks = True
            
            ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명 & ".xlsx"
            ActiveWorkbook.Close savechanges:=True
            
            Application.Wait (Now + TimeValue("0:00:01"))
            ActiveWorkbook.Close

        Next

    Else
        MsgBox "파일을 선택하지 않았습니다."
    End If


End Sub































































'Sub 기지급자료처리()
'
'    Dim 파일형식 As String
'    Dim 대상폴더 As String
'    Dim 파일 As Variant, 선택파일 As Variant
'    Dim 통합시트 As Worksheet
'    Dim 작업파일 As Workbook
'    Dim 대상범위 As Range
'    Dim 복사위치 As Range
'    Dim 행 As Long
'    Dim 행1 As Long
'    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
'    Dim 백업위치 As String
'    Dim 최대행수 As Long
'    Dim 최대열수 As Long
'
'
'
''1. 텍스트 파일 순환가능
'
'
'    파일형식 = "텍스트 파일 (*.txt), *.txt"
'    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업할 텍스트 파일을 선택하세요", MultiSelect:=True)
'
'        '    대상폴더 = ThisWorkbook.Path & "\지급내역\"
'
'    If IsArray(선택파일) = True Then
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'        For Each 파일 In 선택파일
'
'            Workbooks.OpenText Filename:=파일 _
'            , Origin:=949, StartRow:=2, DataType:=xlDelimited, TextQualifier:= _
'            xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, _
'            Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'            Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'            Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2)), TrailingMinusNumbers:=True
'
'
''파일이름은 예전 텍스트에 백업만 붙여서
'
'            현재파일 = ActiveWorkbook.FullName
'            백업위치 = ActiveWorkbook.Path
'            백업파일명 = "백업" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
'
'            ActiveSheet.UsedRange.Select
'            Selection.Copy
'            Workbooks.Add
'
'            최대행수 = ActiveSheet.UsedRange.Rows.Count
'            최대열수 = ActiveSheet.UsedRange.Columns.Count
'
'            For 행 = 1 To 최대행수
'
'                Cells(행, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((행 - 1) \ 4) + 1, 11 * ((행 - 1) Mod 4) + 1)
'
'            Next
'
'            For 행1 = 최대행수 To 1 Step -1
'
'                If 행1 Mod 4 <> 1 Then
'                    Rows(행1).Delete
'                End If
'
'            Next
'
'            Application.ScreenUpdating = True
'            ActiveSheet.DisplayPageBreaks = True
'
'            ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명 & ".txt"
'
'        Next
'
'    Else
'        MsgBox "파일을 선택하지 않았습니다."
'    End If
'
''    If 파일 <> "" Then
''
''        Application.ScreenUpdating = False
''
''            Set 통합시트 = ThisWorkbook.Worksheets(1)
''
''            통합시트.UsedRange.Offset(1).Delete Shift:=xlUp
''
''            Do
''
''                Set 작업파일 = Workbooks.Open(Filename:=대상폴더 & 파일)
''                Set 대상범위 = 작업파일.Worksheets(1).UsedRange
''                With 대상범위
''                    Set 대상범위 = .Offset(1).Resize(.Rows.Count - 1)
''                End With
''
''                Set 복사위치 = 통합시트.Cells(Rows.Count, "A").End(xlUp).Offset(1)
''
''                대상범위.Copy 복사위치
''
''                작업파일.Close SaveChanges:=False
''
''                파일 = Dir()
''
''            Loop While 파일 <> ""
''
''        Application.ScreenUpdating = True
''
''    End If
''
'''2. 텍스트 형식으로 끌고 오기
''
''
''
''    파일형식 = "텍스트 파일 (*.txt), *.txt"
''    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업 파일 선택")
''
''    If 선택파일 <> False Then
''        Workbooks.OpenText Filename:=선택파일 _
''        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
''        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
''        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
''        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
''        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
''        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
''
''    Else
''        MsgBox "파일을 선택하지 않았습니다."
''    End If
''
''    Workbooks.Open (백업위치 & "\" & 백업파일명)
''    Rows("1:11").Delete
''
''    최대행수 = ActiveSheet.UsedRange.Rows.Count
''    최대열수 = ActiveSheet.UsedRange.Columns.Count
''
''    ActiveSheet.DisplayPageBreaks = False
''    Application.ScreenUpdating = False
''
''    For 행 = 1 To 최대행수
''
''        Cells(행, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((행 - 1) \ 4) + 1, 11 * ((행 - 1) Mod 4) + 1)
''
''    Next
''
''    For 행1 = 최대행수 To 1 Step -1
''
''        If 행1 Mod 8 <> 1 Then
''            Rows(행1).Delete
''        End If
''
''    Next
''
''    '빈열 제거
''    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''    Cells.Replace What:="▶▶", Replacement:="", LookAt:=xlWhole, SearchOrder _
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
'Sub 디스켓접수결과처리()
''
'' txt파일불러오기 Macro
''
'
'    Dim 행 As Long
'    Dim 행1 As Long
'    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
'    Dim 백업위치 As String
'    Dim 파일형식 As String
'    Dim 선택파일 As Variant
'    Dim 최대행수 As Long
'    Dim 최대열수 As Long
'
'
'    파일형식 = "텍스트 파일 (*.txt), *.txt"
'    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업 파일 선택")
'
'    If 선택파일 <> False Then
'        Workbooks.OpenText Filename:=선택파일 _
'        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
'        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
'        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
'
'    Else
'        MsgBox "파일을 선택하지 않았습니다."
'    End If
'
'
''DRM풀기
'
'    현재파일 = ActiveWorkbook.FullName
'    백업위치 = ActiveWorkbook.Path
'    백업파일명 = "백업디스켓접수(" & Date & ").xlsx"
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
'    ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명
'
'  '일단 마지막 행 수 구하는 것은 나중에
'
'    Workbooks.Open (백업위치 & "\" & 백업파일명)
'    Rows("1:4").Delete
'
'    최대행수 = ActiveSheet.UsedRange.Rows.Count
'    최대열수 = ActiveSheet.UsedRange.Columns.Count
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'    For 행 = 1 To 최대행수
'
'        Cells(행, 1).Resize(1, 8).Cut Destination:=Cells(13 * ((행 - 1) \ 13) + 1, 8 * ((행 - 1) Mod 13) + 1)
'
'    Next
'
'    For 행1 = 최대행수 To 1 Step -1
'
'        If 행1 Mod 13 <> 1 Then
'            Rows(행1).Delete
'        End If
'
'    Next
'
'    '빈열 제거
'    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="▶▶", Replacement:="", LookAt:=xlWhole, SearchOrder _
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
'    Range("A1") = "환급결정번호"
'    Range("B1") = "가입자주민번호"
'    Range("C1") = "가입자주민번호seq"
'    Range("D1") = "예금주주민번호"
'    Range("E1") = "수진자주민번호"
'    Range("F1") = "수진자주민번호seq"
'    Range("G1") = "최종금액"
'    Range("H1") = "순번"
'    Range("I1") = "은행명"
'    Range("J1") = "가입자성명"
'    Range("K1") = "예금주성명"
'    Range("L1") = "수진자성명"
'    Range("M1") = "증번호"
'    Range("N1") = "계좌번호"
'    Range("O1") = "오류여부"
'    Range("P1") = "저장여부"
'    Range("Q1") = "계좌확인오류"
'    Range("R1") = "계좌명"
'    Range("S1") = "오류유형"
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
''    Range("A1") = "결정직역"
''    Range("B1") = "결정지사"
''    Range("C1") = "결정년월"
''    Range("D1") = "지급종별"
''    Range("E1") = "결정번호"
''    Range("AK1") = "실예금주성명"
''
''    Application.ScreenUpdating = True
''    ActiveSheet.DisplayPageBreaks = True
''
''    Range("F2").Select
''    ActiveWindow.FreezePanes = True
''
''    Selection.AutoFilter
''    ActiveSheet.Range("$A$1:$AL$38").AutoFilter Field:=31, Criteria1:=Array( _
''        "계좌상이", "기접수", "기지급(상태코드)", "="), Operator:=xlFilterValues
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
'Sub 기지급관련작업()
''
'' Macro3 Macro
'' 작업일자 : 2014-10-29
'
''
'    Dim 영역 As Range
'
'    영역 = ActiveSheet.UsedRange
'
'    ActiveWorkbook.Worksheets("기지급계좌통합").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("기지급계좌통합").Sort.SortFields.Add Key:=Range( _
'        "AB2:AB1001"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
'        xlSortTextAsNumbers
'    With ActiveWorkbook.Worksheets("기지급계좌통합").Sort
'        .SetRange 영역
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'    영역.RemoveDuplicates Columns:=Array(6, 24, 25, 26 _
'        , 27), Header:=xlYes
'
'End Sub
'
'
'
'Sub 기존계좌연계결과처리()
''
'' txt파일불러오기 Macro
''
'
'    Dim 행 As Long
'    Dim 행1 As Long
'    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
'    Dim 백업위치 As String
'    Dim 파일형식 As String
'    Dim 선택파일 As Variant
'    Dim 최대행수 As Long
'    Dim 최대열수 As Long
'
'
'    파일형식 = "텍스트 파일 (*.txt), *.txt"
'    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업 파일 선택")
'
'    If 선택파일 <> False Then
'        Workbooks.OpenText Filename:=선택파일 _
'        , Origin:=949, StartRow:=9, DataType:=xlDelimited, TextQualifier:= _
'        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
'        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
'
'    Else
'        MsgBox "파일을 선택하지 않았습니다."
'    End If
'
'
''DRM풀기
'
'    현재파일 = ActiveWorkbook.FullName
'    백업위치 = ActiveWorkbook.Path
'    백업파일명 = "백업" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
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
'    ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명 & ".xlsx"
'
'  '일단 마지막 행 수 구하는 것은 나중에
'
'    Workbooks.Open (백업위치 & "\" & 백업파일명)
'
'    최대행수 = ActiveSheet.UsedRange.Rows.Count
'    최대열수 = ActiveSheet.UsedRange.Columns.Count
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'    For 행 = 1 To 최대행수
'
'        Cells(행, 1).Resize(1, 12).Cut Destination:=Cells(3 * ((행 - 1) \ 3) + 1, 12 * ((행 - 1) Mod 3) + 1)
'
'    Next
'
'    For 행1 = 최대행수 To 1 Step -1
'
'        If 행1 Mod 3 <> 1 Then
'            Rows(행1).Delete
'        End If
'
'    Next
'
'    '빈열 제거
'    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="▶▶", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    Rows("1:1").Select
'    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Selection.EntireColumn.Delete
'
'    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("A1") = "순번"
'    Range("B1") = "자격여부"
'    Range("C1") = "미지급결정번호"
'    Range("D1") = "수진자주민번호"
'    Range("E1") = "수진자명"
'    Range("F1") = "가입자주민번호"
'    Range("G1") = "가입자명"
'    Range("H1") = "미지급금"
'    Range("I1") = "지급안내상태"
'    Range("J1") = "안내통보일"
'    Range("K1") = "저장여부"
'    Range("L1") = "점검결과"
'    Range("M1") = "연계자주민번호"
'    Range("N1") = "연계자명"
'    Range("O1") = "예금주주민번호"
'    Range("P1") = "예금주명"
'    Range("Q1") = "금융기관"
'    Range("R1") = "계좌번호"
'    Range("S1") = "타인여부"
'    Range("T1") = "실예금주명"
'    Range("U1") = "신청자명"
'    Range("V1") = "등록일자"
'
'    최대행수 = ActiveSheet.UsedRange.Rows.Count
'    Rows(최대행수).Delete
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
'Sub 기지급자료계좌연계관련()
'
'    Dim 파일형식 As String
'    Dim 대상폴더 As String
'    Dim 파일 As Variant, 선택파일 As Variant
'    Dim 통합시트 As Worksheet
'    Dim 작업파일 As Workbook
'    Dim 대상범위 As Range
'    Dim 복사위치 As Range
'    Dim 행 As Long
'    Dim 행1 As Long
'    Dim 레코드별행수 As Long
'    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
'    Dim 백업위치 As String
'    Dim 최대행수 As Long
'    Dim 최대열수 As Long
'
'
'
''1. 텍스트 파일 순환가능
'
'
'    파일형식 = "텍스트 파일 (*.txt), *.txt"
'    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업할 텍스트 파일을 선택하세요", MultiSelect:=True)
'
'        '    대상폴더 = ThisWorkbook.Path & "\지급내역\"
'
'    If IsArray(선택파일) = True Then
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'
'        For Each 파일 In 선택파일
'
'            Workbooks.OpenText Filename:=파일 _
'            , Origin:=949, StartRow:=12, DataType:=xlDelimited, TextQualifier:= _
'            xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, _
'            Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'            Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 2), Array(8, 2), _
'            Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2)), TrailingMinusNumbers:=True
'
'
''파일이름은 예전 텍스트에 백업만 붙여서
'
'            현재파일 = ActiveWorkbook.FullName
'            백업위치 = ActiveWorkbook.Path
'            백업파일명 = "백업" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
'
'            ActiveSheet.UsedRange.Select
'            Selection.Copy
'
'            Workbooks.Add
'            Range("A1").Select
'            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
'            최대행수 = ActiveSheet.UsedRange.Rows.Count
'            Rows(최대행수).Delete
'
'            최대행수 = ActiveSheet.UsedRange.Rows.Count
'
'            For 행1 = 최대행수 To 2 Step -1
'
'                레코드별행수 = 행1 Mod 2
'                Select Case 레코드별행수
'
'                    Case 0
'
'                    Cells(행1, 1).Resize(1, 11).Cut Destination:=Cells(2 * ((행1 - 1) \ 2) + 1, 11 * ((행1 - 1) Mod 2) + 1)
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
'            Range("A1") = "순번"
'            Range("B1") = "환급결정번호"
'            Range("C1") = "수진자주민번호"
'            Range("D1") = "수진자주민번호일련번호"
'            Range("E1") = "가입자주민번호"
'            Range("F1") = "가입자주민번호일련번호"
'            Range("G1") = "지급일"
'            Range("H1") = "증번호"
'            Range("I1") = "직역구분"
'            Range("J1") = "금융기관"
'            Range("K1") = "예금주주민번호"
'            Range("L1") = "접수일"
'            Range("M1") = "접수구분"
'            Range("N1") = "수진자성명"
'            Range("O1") = "가입자성명"
'            Range("P1") = "지급상태"
'            Range("Q1") = "지급결정금액"
'            Range("R1") = "계좌번호"
'            Range("S1") = "예금주성명"
'            Range("T1") = "예금주와 관계"
'
'            Application.ScreenUpdating = True
'            ActiveSheet.DisplayPageBreaks = True
'
'            ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명 & ".xlsx"
'            ActiveWorkbook.Close savechanges:=True
'
'            Application.Wait (Now + TimeValue("0:00:03"))
'            ActiveWorkbook.Close
'
'        Next
'
'    Else
'        MsgBox "파일을 선택하지 않았습니다."
'    End If
'
'
'
''    If 파일 <> "" Then
''
''        Application.ScreenUpdating = False
''
''            Set 통합시트 = ThisWorkbook.Worksheets(1)
''
''            통합시트.UsedRange.Offset(1).Delete Shift:=xlUp
''
''            Do
''
''                Set 작업파일 = Workbooks.Open(Filename:=대상폴더 & 파일)
''                Set 대상범위 = 작업파일.Worksheets(1).UsedRange
''                With 대상범위
''                    Set 대상범위 = .Offset(1).Resize(.Rows.Count - 1)
''                End With
''
''                Set 복사위치 = 통합시트.Cells(Rows.Count, "A").End(xlUp).Offset(1)
''
''                대상범위.Copy 복사위치
''
''                작업파일.Close SaveChanges:=False
''
''                파일 = Dir()
''
''            Loop While 파일 <> ""
''
''        Application.ScreenUpdating = True
''
''    End If
''
'''2. 텍스트 형식으로 끌고 오기
''
''
''
''    파일형식 = "텍스트 파일 (*.txt), *.txt"
''    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업 파일 선택")
''
''    If 선택파일 <> False Then
''        Workbooks.OpenText Filename:=선택파일 _
''        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
''        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
''        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
''        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
''        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
''        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
''
''    Else
''        MsgBox "파일을 선택하지 않았습니다."
''    End If
''
''    Workbooks.Open (백업위치 & "\" & 백업파일명)
''    Rows("1:11").Delete
''
''    최대행수 = ActiveSheet.UsedRange.Rows.Count
''    최대열수 = ActiveSheet.UsedRange.Columns.Count
''
''    ActiveSheet.DisplayPageBreaks = False
''    Application.ScreenUpdating = False
''
''    For 행 = 1 To 최대행수
''
''        Cells(행, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((행 - 1) \ 4) + 1, 11 * ((행 - 1) Mod 4) + 1)
''
''    Next
''
''    For 행1 = 최대행수 To 1 Step -1
''
''        If 행1 Mod 8 <> 1 Then
''            Rows(행1).Delete
''        End If
''
''    Next
''
''    '빈열 제거
''    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''    Cells.Replace What:="▶▶", Replacement:="", LookAt:=xlWhole, SearchOrder _
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
'Sub edi결과처리()
''
'' txt파일불러오기 Macro
''
'
'    Dim 행 As Long
'    Dim 행1 As Long
'    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
'    Dim 백업위치 As String
'    Dim 파일형식 As String
'    Dim 선택파일 As Variant
'    Dim 최대행수 As Long
'    Dim 최대열수 As Long
'
'
'    파일형식 = "텍스트 파일 (*.txt), *.txt"
'    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업 파일 선택")
'
'    If 선택파일 <> False Then
'        Workbooks.OpenText Filename:=선택파일 _
'        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
'        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
'        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
'        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
'        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
'
'    Else
'        MsgBox "파일을 선택하지 않았습니다."
'    End If
'
'
''DRM풀기
'
'    현재파일 = ActiveWorkbook.FullName
'    백업위치 = ActiveWorkbook.Path
'    백업파일명 = "백업edi_" & Replace((Range("c5").Value & Range("h5").Value), ".", "") & ".xlsx"
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
'    ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명
'
'  '일단 마지막 행 수 구하는 것은 나중에
'
'    Workbooks.Open (백업위치 & "\" & 백업파일명)
'    Rows("1:11").Delete
'
'    최대행수 = ActiveSheet.UsedRange.Rows.Count
'    최대열수 = ActiveSheet.UsedRange.Columns.Count
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'    For 행 = 1 To 최대행수
'
'        Cells(행, 1).Resize(1, 11).Cut Destination:=Cells(8 * ((행 - 1) \ 8) + 1, 11 * ((행 - 1) Mod 8) + 1)
'
'    Next
'
'    For 행1 = 최대행수 To 1 Step -1
'
'        If 행1 Mod 8 <> 1 Then
'            Rows(행1).Delete
'        End If
'
'    Next
'
'    '빈열 제거
'    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'    Cells.Replace What:="▶▶", Replacement:="", LookAt:=xlWhole, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'    Rows("1:1").Select
'    Selection.SpecialCells(xlCellTypeBlanks).Select
'    Selection.EntireColumn.Delete
'
'    Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
'    Range("A1") = "환급결정번호"
'    Range("B1") = "증번호"
'    Range("C1") = "통합사업장"
'    Range("D1") = "사업장기호"
'    Range("E1") = "수진자성명"
'    Range("F1") = "수진자주민번호"
'    Range("G1") = "수진자주민번호seq"
'    Range("H1") = "작성일자"
'    Range("I1") = "접수일자"
'    Range("J1") = "최종지급금액"
'    Range("K1") = "단위사업장"
'    Range("L1") = "단위사업장세부"
'    Range("M1") = "가입자성명"
'    Range("N1") = "가입자주민번호"
'    Range("O1") = "가입자주민번호seq"
'    Range("P1") = "저장제외"
'    Range("Q1") = "순번"
'    Range("R1") = "시효만료일"
'    Range("S1") = "통보년월"
'    Range("T1") = "은행코드"
'    Range("U1") = "금융기관명"
'    Range("V1") = "예금주성명"
'    Range("W1") = "예금주주민번호"
'    Range("X1") = "계좌번호"
'    Range("Y1") = "사업장전화번호"
'    Range("Z1") = "안내상태"
'    Range("AA1") = "신청처리결과"
'    Range("AB1") = "자료점검상태"
'    Range("AC1") = "EDI종류"
'    Range("AD1") = "금융기관확인"
'    Range("AE1") = "금융기관계좌번호확인"
'    Range("AE1") = "계좌번호확인"
'    Range("AF1") = "예금주주민번호확인"
'    Range("AG1") = "예금주성명확인"
'    Range("AH1") = "실예금주성명"
'    Range("AI1") = "오류메시지"
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
'    Range("A1") = "결정직역"
'    Range("B1") = "결정지사"
'    Range("C1") = "결정년월"
'    Range("D1") = "지급종별"
'    Range("E1") = "결정번호"
'    Range("AK1") = "실예금주성명"
'
'    Application.ScreenUpdating = True
'    ActiveSheet.DisplayPageBreaks = True
'
'    Range("F2").Select
'    ActiveWindow.FreezePanes = True
'
'    Selection.AutoFilter
'    ActiveSheet.Range("$A$1:$AL$38").AutoFilter Field:=31, Criteria1:=Array( _
'        "계좌상이", "기접수", "기지급(상태코드)", "자격제외자", "="), Operator:=xlFilterValues
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
''사업장기호 정렬
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
'Sub 기지급자료처리_최종()
'
'    Dim 파일형식 As String
'    Dim 대상폴더 As String
'    Dim 파일 As Variant, 선택파일 As Variant
'    Dim 통합시트 As Worksheet
'    Dim 작업파일 As Workbook
'    Dim 대상범위 As Range
'    Dim 복사위치 As Range
'    Dim 행 As Long
'    Dim 행1 As Long
'    Dim 레코드별행수 As Long
'    Dim 현재파일 As String, 백업파일 As String, 백업파일명 As String
'    Dim 백업위치 As String
'    Dim 최대행수 As Long
'    Dim 최대열수 As Long
'
'
'
''1. 텍스트 파일 순환가능
'
'
'    파일형식 = "텍스트 파일 (*.txt), *.txt"
'    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업할 텍스트 파일을 선택하세요", MultiSelect:=True)
'
'        '    대상폴더 = ThisWorkbook.Path & "\지급내역\"
'
'    If IsArray(선택파일) = True Then
'
'    ActiveSheet.DisplayPageBreaks = False
'    Application.ScreenUpdating = False
'
'
'        For Each 파일 In 선택파일
'
'            Workbooks.OpenText Filename:=파일 _
'            , Origin:=949, StartRow:=2, DataType:=xlDelimited, TextQualifier:= _
'            xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, Semicolon:=False, _
'            Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
'            Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 1), Array(7, 2), Array(8, 2), _
'            Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2)), TrailingMinusNumbers:=True
'
'
''파일이름은 예전 텍스트에 백업만 붙여서
'
'            현재파일 = ActiveWorkbook.FullName
'            백업위치 = ActiveWorkbook.Path
'            백업파일명 = "백업" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
'
'            ActiveSheet.UsedRange.Select
'            Selection.Copy
'
'            Workbooks.Add
'            Range("A1").Select
'            Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'
'            최대행수 = ActiveSheet.UsedRange.Rows.Count
'            Rows(최대행수).Delete
'            Rows("1:6").Delete
'
'            최대행수 = ActiveSheet.UsedRange.Rows.Count
'
'            For 행1 = 최대행수 To 2 Step -1
'
'                레코드별행수 = 행1 Mod 4
'                Select Case 레코드별행수
'
'                    Case 2
'
'                    Cells(행1, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((행1 - 1) \ 4) + 1, 11 * ((행1 - 1) Mod 4) + 1)
'                    CutCopyMode = False
'
'                    Case 0
'                    Cells(행1, 1).Cut Destination:=Cells(4 * ((행1 - 1) \ 4) + 1, 23)
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
'            ActiveWorkbook.SaveAs Filename:=백업위치 & "\" & 백업파일명 & ".csv", FileFormat:=xlCSV
'            ActiveWorkbook.Close savechanges:=True
'
'            Application.Wait (Now + TimeValue("0:00:03"))
'            ActiveWorkbook.Close
'
'        Next
'
'    Else
'        MsgBox "파일을 선택하지 않았습니다."
'    End If
'
'
'
''    If 파일 <> "" Then
''
''        Application.ScreenUpdating = False
''
''            Set 통합시트 = ThisWorkbook.Worksheets(1)
''
''            통합시트.UsedRange.Offset(1).Delete Shift:=xlUp
''
''            Do
''
''                Set 작업파일 = Workbooks.Open(Filename:=대상폴더 & 파일)
''                Set 대상범위 = 작업파일.Worksheets(1).UsedRange
''                With 대상범위
''                    Set 대상범위 = .Offset(1).Resize(.Rows.Count - 1)
''                End With
''
''                Set 복사위치 = 통합시트.Cells(Rows.Count, "A").End(xlUp).Offset(1)
''
''                대상범위.Copy 복사위치
''
''                작업파일.Close SaveChanges:=False
''
''                파일 = Dir()
''
''            Loop While 파일 <> ""
''
''        Application.ScreenUpdating = True
''
''    End If
''
'''2. 텍스트 형식으로 끌고 오기
''
''
''
''    파일형식 = "텍스트 파일 (*.txt), *.txt"
''    선택파일 = Application.GetOpenFilename(filefilter:=파일형식, Title:="작업 파일 선택")
''
''    If 선택파일 <> False Then
''        Workbooks.OpenText Filename:=선택파일 _
''        , Origin:=949, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
''        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
''        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 2), _
''        Array(2, 2), Array(3, 2), Array(4, 2), Array(5, 2), Array(6, 2), Array(7, 2), Array(8, 2), _
''        Array(9, 2), Array(10, 2), Array(11, 2), Array(12, 2), Array(13, 2), Array(14, 2), Array(15 _
''        , 2), Array(16, 2), Array(17, 2), Array(18, 2)), TrailingMinusNumbers:=True
''
''    Else
''        MsgBox "파일을 선택하지 않았습니다."
''    End If
''
''    Workbooks.Open (백업위치 & "\" & 백업파일명)
''    Rows("1:11").Delete
''
''    최대행수 = ActiveSheet.UsedRange.Rows.Count
''    최대열수 = ActiveSheet.UsedRange.Columns.Count
''
''    ActiveSheet.DisplayPageBreaks = False
''    Application.ScreenUpdating = False
''
''    For 행 = 1 To 최대행수
''
''        Cells(행, 1).Resize(1, 11).Cut Destination:=Cells(4 * ((행 - 1) \ 4) + 1, 11 * ((행 - 1) Mod 4) + 1)
''
''    Next
''
''    For 행1 = 최대행수 To 1 Step -1
''
''        If 행1 Mod 8 <> 1 Then
''            Rows(행1).Delete
''        End If
''
''    Next
''
''    '빈열 제거
''    Cells.Replace What:="-", Replacement:="", LookAt:=xlWhole, SearchOrder _
''        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
''    Cells.Replace What:="▶▶", Replacement:="", LookAt:=xlWhole, SearchOrder _
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
